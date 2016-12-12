VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmOperacaoEspecial 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Operações Especiais"
   ClientHeight    =   6315
   ClientLeft      =   2790
   ClientTop       =   3420
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   9810
   Begin ACTIVESKINLibCtl.SkinLabel lblTotalItem 
      Height          =   180
      Left            =   7935
      OleObjectBlob   =   "OperacaoEspecial.frx":0000
      TabIndex        =   105
      Top             =   2460
      Width           =   705
   End
   Begin ACTIVESKINLibCtl.SkinLabel lblprecoUnit 
      Height          =   195
      Left            =   6255
      OleObjectBlob   =   "OperacaoEspecial.frx":0072
      TabIndex        =   104
      Top             =   2460
      Width           =   1005
   End
   Begin ACTIVESKINLibCtl.SkinLabel lblQuant 
      Height          =   195
      Left            =   5505
      OleObjectBlob   =   "OperacaoEspecial.frx":00EC
      TabIndex        =   103
      Top             =   2460
      Width           =   480
   End
   Begin ACTIVESKINLibCtl.SkinLabel lblDescricao 
      Height          =   195
      Left            =   1230
      OleObjectBlob   =   "OperacaoEspecial.frx":0156
      TabIndex        =   102
      Top             =   2460
      Width           =   720
   End
   Begin ACTIVESKINLibCtl.SkinLabel lblReferencia 
      Height          =   195
      Left            =   210
      OleObjectBlob   =   "OperacaoEspecial.frx":01C6
      TabIndex        =   101
      Top             =   2460
      Width           =   780
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   390
      OleObjectBlob   =   "OperacaoEspecial.frx":0238
      Top             =   5790
   End
   Begin VB.CommandButton cmdLimpa 
      Caption         =   "Limpa"
      Height          =   345
      Left            =   5565
      TabIndex        =   35
      Top             =   5805
      Width           =   1155
   End
   Begin VB.TextBox txtTotalItem 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   315
      Left            =   7905
      Locked          =   -1  'True
      MaxLength       =   17
      TabIndex        =   27
      Top             =   2670
      Width           =   1650
   End
   Begin VB.TextBox txtPrecoUnit 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   315
      Left            =   6240
      MaxLength       =   17
      TabIndex        =   26
      Top             =   2670
      Width           =   1650
   End
   Begin VB.TextBox txtQuant 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   315
      Left            =   5475
      MaxLength       =   5
      TabIndex        =   25
      Top             =   2670
      Width           =   735
   End
   Begin VB.TextBox txtDescricao 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   315
      Left            =   1200
      MaxLength       =   38
      TabIndex        =   24
      Top             =   2670
      Width           =   4275
   End
   Begin VB.TextBox txtReferencia 
      BackColor       =   &H00C0C0C0&
      Height          =   315
      Left            =   195
      MaxLength       =   7
      TabIndex        =   23
      Top             =   2670
      Width           =   975
   End
   Begin VB.CommandButton cmdGrava 
      Caption         =   "&Grava"
      Enabled         =   0   'False
      Height          =   345
      Left            =   6735
      TabIndex        =   34
      Top             =   5805
      Width           =   1395
   End
   Begin VB.CommandButton cmdRetorna 
      Caption         =   "&Retorna"
      Height          =   345
      Left            =   8145
      TabIndex        =   36
      Top             =   5805
      Width           =   1395
   End
   Begin VB.Frame fraDadosNota 
      Height          =   2445
      Left            =   45
      TabIndex        =   38
      Top             =   -30
      Width           =   9660
      Begin ACTIVESKINLibCtl.SkinLabel lblCEP 
         Height          =   195
         Left            =   8415
         OleObjectBlob   =   "OperacaoEspecial.frx":046C
         TabIndex        =   100
         Top             =   1800
         Width           =   315
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblUFCli 
         Height          =   195
         Left            =   7815
         OleObjectBlob   =   "OperacaoEspecial.frx":04D0
         TabIndex        =   99
         Top             =   1800
         Width           =   210
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   195
         Left            =   6375
         OleObjectBlob   =   "OperacaoEspecial.frx":0532
         TabIndex        =   98
         Top             =   1800
         Width           =   630
      End
      Begin VB.Frame frmCliente 
         Caption         =   "Tipo de Cliente"
         ForeColor       =   &H000000FF&
         Height          =   780
         Left            =   3300
         TabIndex        =   75
         Top             =   1425
         Visible         =   0   'False
         Width           =   2610
         Begin VB.OptionButton OptCliente 
            Caption         =   "Jurídico"
            ForeColor       =   &H00FF0000&
            Height          =   270
            Index           =   1
            Left            =   1440
            TabIndex        =   77
            Top             =   330
            Width           =   1050
         End
         Begin VB.OptionButton OptCliente 
            Caption         =   "Físico"
            ForeColor       =   &H00FF0000&
            Height          =   270
            Index           =   0
            Left            =   240
            TabIndex        =   76
            Top             =   330
            Width           =   1050
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblMunicipio 
         Height          =   195
         Left            =   3285
         OleObjectBlob   =   "OperacaoEspecial.frx":05A0
         TabIndex        =   97
         Top             =   1800
         Width           =   705
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblBairro 
         Height          =   195
         Left            =   120
         OleObjectBlob   =   "OperacaoEspecial.frx":0610
         TabIndex        =   96
         Top             =   1800
         Width           =   405
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblRazao 
         Height          =   195
         Left            =   120
         OleObjectBlob   =   "OperacaoEspecial.frx":067A
         TabIndex        =   94
         Top             =   1260
         Width           =   465
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblInscEst 
         Height          =   195
         Left            =   7965
         OleObjectBlob   =   "OperacaoEspecial.frx":06E2
         TabIndex        =   93
         Top             =   720
         Width           =   1305
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblCGCTransf 
         Height          =   195
         Left            =   6450
         OleObjectBlob   =   "OperacaoEspecial.frx":0764
         TabIndex        =   92
         Top             =   720
         Width           =   1350
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblCliente 
         Height          =   195
         Left            =   5625
         OleObjectBlob   =   "OperacaoEspecial.frx":07E4
         TabIndex        =   91
         Top             =   720
         Width           =   480
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblDataPagto 
         Height          =   195
         Left            =   4410
         OleObjectBlob   =   "OperacaoEspecial.frx":0850
         TabIndex        =   89
         Top             =   735
         Visible         =   0   'False
         Width           =   810
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblCondPagto 
         Height          =   195
         Left            =   120
         OleObjectBlob   =   "OperacaoEspecial.frx":08C2
         TabIndex        =   87
         Top             =   735
         Width           =   1530
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblDataEmi 
         Height          =   195
         Left            =   8310
         OleObjectBlob   =   "OperacaoEspecial.frx":0944
         TabIndex        =   86
         Top             =   180
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblVendLojaVenda 
         Height          =   195
         Left            =   7440
         OleObjectBlob   =   "OperacaoEspecial.frx":09BA
         TabIndex        =   85
         Top             =   180
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblVendedor 
         Height          =   195
         Left            =   6675
         OleObjectBlob   =   "OperacaoEspecial.frx":0A2E
         TabIndex        =   84
         Top             =   180
         Width           =   690
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblLjDestino 
         Height          =   195
         Left            =   5955
         OleObjectBlob   =   "OperacaoEspecial.frx":0A9C
         TabIndex        =   83
         Top             =   180
         Width           =   540
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblNatOper 
         Height          =   195
         Left            =   3090
         OleObjectBlob   =   "OperacaoEspecial.frx":0B0A
         TabIndex        =   82
         Top             =   180
         Width           =   1395
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblCFO 
         Height          =   195
         Left            =   2385
         OleObjectBlob   =   "OperacaoEspecial.frx":0B8A
         TabIndex        =   81
         Top             =   180
         Width           =   315
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblLjOrigem 
         Height          =   195
         Left            =   1635
         OleObjectBlob   =   "OperacaoEspecial.frx":0BEE
         TabIndex        =   80
         Top             =   195
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblSerie 
         Height          =   195
         Left            =   1110
         OleObjectBlob   =   "OperacaoEspecial.frx":0C5C
         TabIndex        =   79
         Top             =   180
         Width           =   360
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblNF 
         Height          =   165
         Left            =   150
         OleObjectBlob   =   "OperacaoEspecial.frx":0CC4
         TabIndex        =   78
         Top             =   180
         Width           =   825
      End
      Begin VB.TextBox txtFoneCli 
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   315
         Left            =   6360
         MaxLength       =   40
         TabIndex        =   20
         Top             =   2010
         Width           =   1410
      End
      Begin VB.ComboBox cmbNaturezaOperacao 
         BackColor       =   &H80000004&
         Height          =   315
         Left            =   3075
         TabIndex        =   4
         Text            =   "cmbNaturezaOperacao"
         Top             =   390
         Width           =   2775
      End
      Begin VB.TextBox txtOutroVendedor 
         BackColor       =   &H8000000A&
         Height          =   315
         Left            =   7470
         MaxLength       =   20
         TabIndex        =   7
         Top             =   390
         Width           =   765
      End
      Begin VB.TextBox txtAV 
         BackColor       =   &H80000004&
         Height          =   315
         Left            =   3645
         TabIndex        =   10
         Top             =   930
         Width           =   765
      End
      Begin VB.ComboBox cmbUFCli 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   7800
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   2010
         Width           =   600
      End
      Begin MSMask.MaskEdBox mskDataPagto 
         Height          =   315
         Left            =   4440
         TabIndex        =   11
         Top             =   930
         Visible         =   0   'False
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483644
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDataEmi 
         Height          =   315
         Left            =   8280
         TabIndex        =   8
         Top             =   405
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483644
         Enabled         =   0   'False
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtCEP 
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   315
         Left            =   8415
         MaxLength       =   8
         TabIndex        =   22
         Top             =   2010
         Width           =   1065
      End
      Begin VB.TextBox txtPagtoEnt 
         BackColor       =   &H8000000A&
         Height          =   315
         Left            =   4440
         MaxLength       =   17
         TabIndex        =   12
         Top             =   930
         Width           =   1140
      End
      Begin VB.TextBox txtCGC 
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   315
         Left            =   6405
         MaxLength       =   15
         TabIndex        =   14
         Top             =   930
         Width           =   1515
      End
      Begin VB.TextBox txtCliente 
         BackColor       =   &H8000000A&
         Height          =   315
         Left            =   5595
         MaxLength       =   6
         TabIndex        =   13
         Top             =   930
         Width           =   780
      End
      Begin VB.TextBox txtLjDestino 
         BackColor       =   &H8000000A&
         Height          =   315
         Left            =   5895
         MaxLength       =   5
         TabIndex        =   5
         Top             =   390
         Width           =   720
      End
      Begin VB.TextBox txtVendedor 
         BackColor       =   &H8000000A&
         Height          =   315
         Left            =   6660
         MaxLength       =   20
         TabIndex        =   6
         Top             =   390
         Width           =   765
      End
      Begin VB.TextBox txtSerie 
         BackColor       =   &H8000000A&
         Height          =   315
         Left            =   1095
         MaxLength       =   2
         TabIndex        =   1
         Top             =   390
         Width           =   450
      End
      Begin VB.TextBox txtNF 
         BackColor       =   &H8000000A&
         Height          =   315
         Left            =   120
         MaxLength       =   6
         TabIndex        =   0
         Top             =   390
         Width           =   885
      End
      Begin VB.TextBox txtCFO 
         BackColor       =   &H8000000A&
         Height          =   315
         Left            =   2370
         MaxLength       =   4
         TabIndex        =   3
         Top             =   390
         Width           =   645
      End
      Begin VB.TextBox txtLjOrigem 
         BackColor       =   &H8000000A&
         Height          =   315
         Left            =   1605
         MaxLength       =   5
         TabIndex        =   2
         Top             =   390
         Width           =   720
      End
      Begin VB.ComboBox cmbCondPagto 
         BackColor       =   &H80000004&
         Height          =   315
         Left            =   90
         TabIndex        =   9
         Top             =   930
         Width           =   3510
      End
      Begin VB.TextBox txtInscEst 
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   315
         Left            =   7950
         MaxLength       =   15
         TabIndex        =   15
         Top             =   930
         Width           =   1530
      End
      Begin VB.TextBox txtRazao 
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   315
         Left            =   90
         MaxLength       =   40
         TabIndex        =   16
         Top             =   1470
         Width           =   5250
      End
      Begin VB.TextBox txtEndereco 
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   315
         Left            =   5415
         MaxLength       =   40
         TabIndex        =   17
         Top             =   1455
         Width           =   4065
      End
      Begin VB.TextBox txtBairro 
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   315
         Left            =   90
         MaxLength       =   15
         TabIndex        =   18
         Top             =   2010
         Width           =   3150
      End
      Begin VB.TextBox txtMunicipio 
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   315
         Left            =   3285
         MaxLength       =   15
         TabIndex        =   19
         Top             =   2010
         Width           =   3030
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblAV 
         Height          =   195
         Left            =   3675
         OleObjectBlob   =   "OperacaoEspecial.frx":0D38
         TabIndex        =   88
         Top             =   750
         Width           =   210
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblPagtoEnt 
         Height          =   225
         Left            =   4410
         OleObjectBlob   =   "OperacaoEspecial.frx":0D9A
         TabIndex        =   90
         Top             =   735
         Width           =   1035
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblEndereco 
         Height          =   195
         Left            =   5430
         OleObjectBlob   =   "OperacaoEspecial.frx":0E12
         TabIndex        =   95
         Top             =   1245
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telefone"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   6375
         TabIndex        =   74
         Top             =   1800
         Width           =   630
      End
      Begin VB.Label lblVendLojaVenda2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Outro Vend."
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   7440
         TabIndex        =   73
         Top             =   180
         Width           =   855
      End
      Begin VB.Label lblAV2 
         AutoSize        =   -1  'True
         Caption         =   "AV"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3675
         TabIndex        =   72
         Top             =   750
         Width           =   210
      End
      Begin VB.Label lblDataPagto2 
         AutoSize        =   -1  'True
         Caption         =   "Data Pagto"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4410
         TabIndex        =   70
         Top             =   735
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label lblCGCTransf2 
         AutoSize        =   -1  'True
         Caption         =   "CGC Transferência"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   6450
         TabIndex        =   69
         Top             =   720
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.Label lblCEP2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CEP"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   8415
         TabIndex        =   62
         Top             =   1800
         Width           =   315
      End
      Begin VB.Label lblUFCli2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UF"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   7815
         TabIndex        =   61
         Top             =   1800
         Width           =   210
      End
      Begin VB.Label lblMunicipio2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Município"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3285
         TabIndex        =   60
         Top             =   1800
         Width           =   705
      End
      Begin VB.Label lblBairro2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   59
         Top             =   1800
         Width           =   405
      End
      Begin VB.Label lblEndereco2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5430
         TabIndex        =   58
         Top             =   1245
         Width           =   690
      End
      Begin VB.Label lblSerie2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Série"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1110
         TabIndex        =   52
         Top             =   180
         Width           =   360
      End
      Begin VB.Label lblCFO2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CFO"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2385
         TabIndex        =   51
         Top             =   180
         Width           =   315
      End
      Begin VB.Label lblLjDestino2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lj Dest."
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5955
         TabIndex        =   50
         Top             =   180
         Width           =   540
      End
      Begin VB.Label lblDataEmi2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Emissão"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   8295
         TabIndex        =   49
         Top             =   180
         Width           =   975
      End
      Begin VB.Label lblVendedor2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   6675
         TabIndex        =   48
         Top             =   180
         Width           =   690
      End
      Begin VB.Label lblCliente2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5625
         TabIndex        =   47
         Top             =   720
         Width           =   480
      End
      Begin VB.Label lblPagtoEnt2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pagto Entrada"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4410
         TabIndex        =   46
         Top             =   735
         Width           =   1035
      End
      Begin VB.Label lblCondPagto2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Condição Pagamento"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   735
         Width           =   1530
      End
      Begin VB.Label lblNF2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nota Fiscal"
         ForeColor       =   &H00FF0000&
         Height          =   165
         Left            =   150
         TabIndex        =   44
         Top             =   180
         Width           =   825
      End
      Begin VB.Label lblLjOrigem2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lj Orig."
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1635
         TabIndex        =   43
         Top             =   195
         Width           =   495
      End
      Begin VB.Label lblNatOper2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Natureza Operação"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3090
         TabIndex        =   42
         Top             =   180
         Width           =   1395
      End
      Begin VB.Label lblInscEst2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inscrição Estadual"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   7965
         TabIndex        =   41
         Top             =   720
         Width           =   1305
      End
      Begin VB.Label lblRazao2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Razão"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   1260
         Width           =   465
      End
   End
   Begin VB.Frame fraDadosTotal 
      Height          =   960
      Left            =   195
      TabIndex        =   39
      Top             =   4650
      Width           =   9375
      Begin ACTIVESKINLibCtl.SkinLabel lblTotal 
         Height          =   195
         Left            =   6435
         OleObjectBlob   =   "OperacaoEspecial.frx":0E80
         TabIndex        =   111
         Top             =   585
         Width           =   360
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblFrete 
         Height          =   195
         Left            =   6435
         OleObjectBlob   =   "OperacaoEspecial.frx":0EE8
         TabIndex        =   110
         Top             =   225
         Width           =   360
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblValorIPI 
         Height          =   195
         Left            =   3240
         OleObjectBlob   =   "OperacaoEspecial.frx":0F50
         TabIndex        =   109
         Top             =   585
         Width           =   600
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblAliqIPI 
         Height          =   195
         Left            =   3210
         OleObjectBlob   =   "OperacaoEspecial.frx":0FC0
         TabIndex        =   108
         Top             =   240
         Width           =   300
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblDesconto 
         Height          =   195
         Left            =   105
         OleObjectBlob   =   "OperacaoEspecial.frx":1026
         TabIndex        =   107
         Top             =   585
         Width           =   690
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblSubTotal 
         Height          =   195
         Left            =   105
         OleObjectBlob   =   "OperacaoEspecial.frx":1094
         TabIndex        =   106
         Top             =   225
         Width           =   690
      End
      Begin VB.TextBox txtTotal 
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   315
         Left            =   7515
         MaxLength       =   17
         TabIndex        =   33
         Top             =   480
         Width           =   1755
      End
      Begin VB.TextBox txtFrete 
         BackColor       =   &H8000000A&
         Height          =   315
         Left            =   7515
         MaxLength       =   17
         TabIndex        =   32
         Top             =   135
         Width           =   1755
      End
      Begin VB.TextBox txtValorIPI 
         BackColor       =   &H8000000A&
         Height          =   315
         Left            =   4395
         MaxLength       =   17
         TabIndex        =   31
         Top             =   495
         Width           =   1755
      End
      Begin VB.TextBox txtAliqIPI 
         BackColor       =   &H8000000A&
         Height          =   315
         Left            =   4395
         MaxLength       =   5
         TabIndex        =   30
         Top             =   135
         Width           =   645
      End
      Begin VB.TextBox txtDesconto 
         BackColor       =   &H8000000A&
         Height          =   315
         Left            =   1230
         MaxLength       =   17
         TabIndex        =   29
         Top             =   480
         Width           =   1755
      End
      Begin VB.TextBox txtSubTotal 
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1230
         MaxLength       =   17
         TabIndex        =   28
         Top             =   135
         Width           =   1755
      End
      Begin VB.Label lblFrete2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Frete"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   6435
         TabIndex        =   68
         Top             =   225
         Width           =   360
      End
      Begin VB.Label lblAliqIPI2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%IPI"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3210
         TabIndex        =   67
         Top             =   240
         Width           =   300
      End
      Begin VB.Label lblValorIPI2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor IPI"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3240
         TabIndex        =   66
         Top             =   585
         Width           =   600
      End
      Begin VB.Label lblTotal2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   6435
         TabIndex        =   65
         Top             =   585
         Width           =   360
      End
      Begin VB.Label lblDesconto2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   105
         TabIndex        =   64
         Top             =   585
         Width           =   690
      End
      Begin VB.Label lblSubTotal2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sub-Total"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   105
         TabIndex        =   63
         Top             =   225
         Width           =   690
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdItens 
      Height          =   1635
      Left            =   180
      TabIndex        =   37
      Top             =   3015
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   2884
      _Version        =   393216
      Cols            =   5
      ForeColorFixed  =   192
      GridColorFixed  =   64
      FormatString    =   "<Referência|<Descrição|>Quantidade|>Preço Unitário|>Total Item"
   End
   Begin VB.Label lblObservacao 
      AutoSize        =   -1  'True
      Caption         =   "Observação"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1620
      TabIndex        =   71
      Top             =   4695
      Width           =   870
   End
   Begin VB.Label lblReferencia2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Referência"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   210
      TabIndex        =   57
      Top             =   2460
      Width           =   780
   End
   Begin VB.Label lblprecoUnit2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preço Unitário"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   6255
      TabIndex        =   56
      Top             =   2460
      Width           =   1005
   End
   Begin VB.Label lblQuant2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quant."
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   5505
      TabIndex        =   55
      Top             =   2460
      Width           =   480
   End
   Begin VB.Label lblDescricao2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1230
      TabIndex        =   54
      Top             =   2460
      Width           =   720
   End
   Begin VB.Label lblTotalItem2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Item"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   7935
      TabIndex        =   53
      Top             =   2460
      Width           =   705
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   165
      X2              =   9555
      Y1              =   5730
      Y2              =   5730
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      Index           =   0
      X1              =   165
      X2              =   9555
      Y1              =   5715
      Y2              =   5715
   End
End
Attribute VB_Name = "frmOperacaoEspecial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Type Itens
   ValorIPI As Double
   Desconto As Double
End Type

Private Type Natureza
    CFO As Long
    Descricao As String
End Type

Dim matNatureza() As Natureza

Dim AlteraDados As Boolean
Dim ClienteExiste As Boolean
Dim EmiteNota As Boolean
Dim GravacaoOK As Boolean
Dim ParaImpr As Boolean
Dim CodigoOperAntigo As Integer
Dim CodigoOperNovo As Integer

Dim MatItens() As Itens


Dim Linhas As Long
Dim Altera As Long

Dim VdCp As String

Dim rdoNumReq As rdoResultset
Dim rdoProduto As rdoResultset
Dim rdoNumeroNF As rdoResultset
Dim rdoVendedor As rdoResultset
Dim rdoConsuNFobs As rdoResultset
Dim RdoCarimbo As rdoResultset
Dim rdoCombo As rdoResultset
Dim RdoNota As rdoResultset
Dim Conexao As New rdoConnection

Dim RSPegaProduto As rdoResultset
Dim RSRegiao As rdoResultset
Dim RSNota As rdoResultset
Dim ClienteDBF As rdoResultset

Dim wVlDescRat As Double
Dim SubTotal As Double
Dim Desconto As Double
Dim TotalGeral As Double
Dim Peso As Double
Dim SomaPeso As Double
Dim Numeronf As Double
Dim NumReq As Double

Private Sub cmbCondPagto_LostFocus()

    If mskDataEmi <> "__/__/____" And mskDataEmi = Wdata Then
        EncontraTexto cmbCondPagto
        
        If cmbCondPagto.Text = "" Then
           cmbCondPagto.ListIndex = 0
        End If
        
        If Val(cmbCondPagto.Text) = 85 Then
           lblPagtoEnt.Visible = False
           txtPagtoEnt.Visible = False
           lblDataPagto.Visible = True
           mskDataPagto.Visible = True
           Faturada = True
           txtAV.SetFocus
        ElseIf Val(cmbCondPagto.Text) > 3 Then
           lblDataPagto.Visible = False
           mskDataPagto.Visible = False
           lblPagtoEnt.Visible = True
           txtPagtoEnt.Visible = True
           txtPagtoEnt.Enabled = True
           Faturada = True
           txtAV.SetFocus
        ElseIf Val(cmbCondPagto.Text) = 3 Then
           lblDataPagto.Visible = False
           mskDataPagto.Visible = False
           lblPagtoEnt.Visible = True
           txtPagtoEnt.Visible = True
           txtPagtoEnt.Enabled = True
           Financiada = True
           txtAV.SetFocus
        Else
           lblDataPagto.Visible = False
           mskDataPagto.Visible = False
           lblPagtoEnt.Visible = True
           txtPagtoEnt.Visible = True
           txtAV.Enabled = True
           txtAV.SetFocus
        End If
    End If
End Sub

Private Sub cmbNaturezaOperacao_GotFocus()

    cmbNaturezaOperacao.ListIndex = 0

End Sub

Private Sub cmbNaturezaOperacao_LostFocus()

   On Error Resume Next
   
   If cmbNaturezaOperacao.Text = "" Then
      cmbNaturezaOperacao.ListIndex = 0
   End If
   
   If txtCFO.Text = 132 Or txtCFO.Text = 232 Then
      txtSerie.Text = "SN"
      txtSerie.Enabled = False
   End If
   
End Sub

Private Sub cmdGrava_Click()
   
   Dim Desconto As Double
   Dim ValorIPI As Double
   Dim AcumulaDesc As Double
   Dim DifDesc As Double
   Dim AcumulaIPI As Double
   Dim DifIPI As Double
   
   GravacaoOK = False
   
   If Val(txtCFO.Text) > 0 Then
      Screen.MousePointer = 11
      
      Linhas = grdItens.Rows - 1
      
      ReDim MatItens(1 To Linhas) As Itens
      
      AcumulaDesc = 0
      
      If Trim(txtDesconto.Text) = "" Then
         txtDesconto.Text = "0,00"
      End If
      
      If Trim(txtValorIPI.Text) = "" Then
         txtValorIPI.Text = "0,00"
      End If
      
      If Trim(txtFrete.Text) = "" Then
         txtFrete.Text = "0,00"
      End If
         
      'If Val(txtCliente.Text) <> 0 Or txtCliente.Text = "" Then
      If Val(txtCliente.Text) <> 0 Or txtCliente.Text <> "" Then
         If cmbUFCli.Text = "" And Val(txtCliente.Text) <> 999999 Then
            Screen.MousePointer = 0
            MsgBox "Favor preencher o campo UF", vbInformation, "Informação"
            cmbUFCli.Enabled = True
            cmbUFCli.SetFocus
            Exit Sub
         End If
      End If
      
      If Val(txtCFO.Text) <> 5152 Then
         If grdItens.Rows > 1 Then
            For Linhas = 1 To grdItens.Rows - 2
               Desconto = CDbl(Format(grdItens.TextMatrix(Linhas, 4), "0.00")) * CDbl(Format(txtDesconto.Text, "0.00")) / CDbl(Format(txtSubTotal.Text, "0.00"))
               AcumulaDesc = AcumulaDesc + Desconto
               DifDesc = CDbl(Format(txtDesconto.Text, "0.00")) - AcumulaDesc
               MatItens(Linhas).Desconto = Desconto
               
               If txtAliqIPI.Text = "" Then
                  ValorIPI = "0,00"
                  DifIPI = -AcumulaIPI
               Else
                  ValorIPI = CDbl(Format(txtAliqIPI.Text, "0.00")) * CDbl(Format(grdItens.TextMatrix(Linhas, 3), "0.00")) * CDbl(Format(grdItens.TextMatrix(Linhas, 2), "0.00"))
                  AcumulaIPI = AcumulaIPI + ValorIPI
                  DifIPI = CDbl(Format(txtAliqIPI.Text, "0.00")) - AcumulaIPI
               End If
               
               MatItens(Linhas).ValorIPI = ValorIPI
            Next Linhas
         
            MatItens(grdItens.Rows - 1).Desconto = DifDesc
            MatItens(grdItens.Rows - 1).ValorIPI = DifIPI
         
         Else
            MatItens(grdItens.Rows - 1).Desconto = CDbl(Format(txtDesconto.Text, "0.00"))
            MatItens(grdItens.Rows - 1).ValorIPI = CDbl(Format(txtAliqIPI.Text, "0.00"))
         End If
      End If
      
      If txtTotal.Text = "" Then
         txtTotal.Text = "0.00"
      End If
      
      If DateDiff("d", Format(mskDataEmi.Text, "dd/mm/yyyy"), Format(Date, "dd/mm/yyyy")) >= 0 Then
         
         If txtCFO.Text <> 5152 And txtCFO.Text <> 5202 And txtCFO.Text <> 6202 Then
            TotalGeral = CDbl(txtSubTotal.Text) - CDbl(txtDesconto.Text) + CDbl(txtValorIPI.Text) + CDbl(txtFrete.Text)
                  
            If TotalGeral <> 0 Then
               If TotalGeral <> CDbl(txtTotal.Text) Then
                  If MsgBox("Soma do total é diferente do total digitado. Confirma? ", vbYesNo + vbQuestion + vbDefaultButton2, "Digitação Notas Manuais") = vbYes Then
                     If AlteraDados = False Then
                        Gravar
                     Else
                        Alteracao
                     End If
                  Else
                     Screen.MousePointer = 0
                     txtTotal.SetFocus
                     Exit Sub
                  End If
                  
               Else
                  If AlteraDados = True Then
                     Alteracao
                  Else
                     Gravar
                  End If
               End If
            End If
         Else
            
            TotalGeral = (txtSubTotal.Text - txtDesconto.Text + txtValorIPI.Text + txtFrete.Text)
            
            If TotalGeral <> txtTotal.Text Then
               If MsgBox("Soma do total é diferente do total digitado. Confirma? ", vbYesNo + vbQuestion + vbDefaultButton2, "Digitação Notas Manuais") = vbYes Then
                  If AlteraDados = False Then
                     Gravar
                  Else
                     Alteracao
                  End If
               Else
                  Screen.MousePointer = 0
                  txtTotal.SetFocus
                  Exit Sub
               End If
                  
            Else
               If AlteraDados = True Then
                  Alteracao
               Else
                  Gravar
               End If
            End If
               
         End If
      Else
         Screen.MousePointer = 0
         MsgBox "Data de emissão superior à data do sistema", vbCritical, "Atenção"
         mskDataEmi.SelStart = 0
         mskDataEmi.SelLength = Len(mskDataEmi.Text)
         mskDataEmi.SelStart = 0
         mskDataEmi.SelLength = Len(mskDataEmi.Text)
         mskDataEmi.SetFocus
         Exit Sub
      End If
         
      Screen.MousePointer = 0
      
   End If
   If GravacaoOK = True Then
      
      EncerraVenda Val(WnumeroPed), txtSerie, 0
      
      If txtCFO.Text = "5152" Then
         SQL = ""
         SQL = "Select * From CTCaixa Where CT_Data = '" & Format(Wdata, "mm/dd/yyyy") & "' "
         Set ISQL = rdoCnLojaBach.OpenResultset(SQL)
         
         rdoCnLojaBach.Execute "Insert into movimentocaixa (MC_NumeroEcf,MC_CodigoOperador,MC_Loja,MC_Data,MC_Grupo,MC_Documento,MC_Serie,MC_Valor,MC_banco,MC_Agencia," _
                     & "MC_Contacorrente,MC_bomPara,MC_Parcelas, MC_Remessa,MC_SituacaoEnvio) values(" & glb_ECF & ",'" & ISQL("ct_operador") & "','" & ISQL("ct_loja") & "', " _
                     & " '" & Format(ISQL("ct_data"), "mm/dd/yyyy") & "', " & 20109 & "," & txtNF.Text & ",'" & txtSerie.Text & "' , " _
                     & "" & ConverteVirgula(Format(txtTotal.Text, "###,###0.00")) & ", " _
                     & "0,0,0,0,0, " & 9 & ",'A')"
      ElseIf WTipoNota <> "S" And WTipoNota <> "E" Then
          frmModalidadeNotaManual.Show 1
      Else
          AtualizaEstoque txtNF.Text, txtSerie.Text, 1
      End If
      fraDadosTotal.Visible = True
      SubTotal = 0
      txtNF.SetFocus
      LimpaTela
   End If
End Sub

Private Sub cmdLimpa_Click()
   LimpaTela
   AlteraDados = False
   fraDadosTotal.Visible = True
   Altera = 0
End Sub

Private Sub cmdRetorna_Click()
   Conexao.Close
   Unload Me
End Sub

Private Sub Form_Load()
   
    Skin1.LoadSkin App.Path & "\Skin\Mxs3.skn"
    Skin1.ApplySkin Me.hwnd
    
    frmOperacaoEspecial.Top = (Screen.Height - frmOperacaoEspecial.Height) / 4
    frmOperacaoEspecial.Left = (Screen.Width - frmOperacaoEspecial.Width) / 2
   
   
   Dim VendaCompra As String
   Dim rdoCombo As rdoResultset
   Dim CondPagto As String
   Dim Descricao As String
   
   AlteraDados = False
   
   mskDataPagto.Text = "__/__/____"
   txtPagtoEnt.Text = ""
   lblCGCTransf.Visible = True
   cmbNaturezaOperacao.Text = ""
   
   grdItens.ColWidth(0) = 1000
   grdItens.ColWidth(1) = 4000
   grdItens.ColWidth(2) = 900
   grdItens.ColWidth(3) = 1300
   grdItens.ColWidth(4) = 1300
      
   SubTotal = 0
   
   If Not ConectaODBC(Conexao, GLB_Usuario, GLB_Senha) Then
        Set Conexao = Nothing
        
        MsgBox "Não foi possível efetuar conexão ao Banco de Dados. Tente feche a tela e tente novamente!", vbCritical, "Erro"
        
        Unload Me
   End If
      
   CarregaNaturezaOperacao

   PreencheComboUF cmbUFCli
   
   Call ExtraiDataMovimento
   
   mskDataEmi.Text = Format(Wdata, "dd/mm/yyyy")
   
   mskDataEmi.Enabled = False
   
End Sub



Private Sub grdItens_DblClick()
   If grdItens.Rows <= 2 And grdItens.TextMatrix(grdItens.Row, 0) = "" Then
      MsgBox "Atenção não há itens", vbInformation, "Atençao"
      Exit Sub
   End If
    
    
   Altera = 1
   
   txtreferencia.Text = grdItens.TextMatrix(grdItens.Row, 0)
   txtdescricao.Text = grdItens.TextMatrix(grdItens.Row, 1)
   txtQuant.Text = grdItens.TextMatrix(grdItens.Row, 2)
   txtPrecoUnit.Text = grdItens.TextMatrix(grdItens.Row, 3)
   txtTotalItem.Text = grdItens.TextMatrix(grdItens.Row, 4)
   txtQuant.Enabled = True
   txtPrecoUnit.Enabled = True
   If Altera = 1 Then
      SubTotal = txtSubTotal.Text
      SubTotal = SubTotal - txtTotalItem
   End If
   
   txtreferencia.SetFocus
   
End Sub

Private Sub grdItens_KeyDown(keycode As Integer, Shift As Integer)

   If keycode = vbKeyDelete Then
      If grdItens.Rows = 2 Then
         txtSubTotal.Text = (txtSubTotal.Text) - (grdItens.TextMatrix(grdItens.Row, 4))
         txtTotal.Text = (Val(txtTotal.Text) + Val(txtDesconto.Text)) - (grdItens.TextMatrix(grdItens.Row, 4))
         grdItens.AddItem ""
         grdItens.RemoveItem (grdItens.Row)
         cmdGrava.Enabled = False
         txtDesconto.Text = ""
         txtAliqIPI.Text = ""
         txtValorIPI.Text = ""
         txtFrete.Text = ""
         Desconto = 0
         txtQuant.Enabled = False
         txtPrecoUnit.Enabled = False
      Else
         txtSubTotal.Text = (txtSubTotal.Text) - (grdItens.TextMatrix(grdItens.Row, 4))
         txtTotal.Text = (txtTotal.Text) - (grdItens.TextMatrix(grdItens.Row, 4))
         grdItens.RemoveItem grdItens.Row
      End If
      SubTotal = txtSubTotal.Text
   End If
   
End Sub

Private Sub mskDataEmi_LostFocus()
    If mskDataEmi <> "__/__/____" Then
        If GetAsyncKeyState(vbKeyTab) <> 0 Then
             Call ExtraiDataMovimento
             If Format(mskDataEmi, "dd/mm/yyyy") <> Wdata Then
                 MsgBox "Data não permitida", vbInformation, "Atenção"
                 mskDataEmi.SelStart = 0
                 mskDataEmi.SelLength = Len(mskDataEmi)
                 mskDataEmi.SetFocus
             End If
         End If
    Else
        MsgBox "Favor preencher a data de emissão", vbInformation, "Atenção"
        mskDataEmi.SelStart = 0
        mskDataEmi.SelLength = Len(mskDataEmi)
        mskDataEmi.SetFocus
    End If
End Sub

Private Sub mskDataPagto_LostFocus()

   If Left(cmbCondPagto.Text, 2) = 85 Then
      If mskDataPagto.Text = "__/__/____" Then
         MsgBox "Favor preencher a data de pagamento", vbInformation, "Informação!"
         mskDataPagto.SelStart = 0
         mskDataPagto.SelLength = Len(mskDataPagto.Text)
         mskDataPagto.SetFocus
      End If
   End If
   
End Sub

Private Sub OptCliente_Click(Index As Integer)

    If OptCliente(0).Value = True Then
        frmCliente.Visible = False
        txtCGC.Enabled = True
        txtCGC.SetFocus
        wPessoa = 2
        OptCliente(0).Value = False
    ElseIf OptCliente(1).Value = True Then
        frmCliente.Visible = False
        txtCGC.Enabled = True
        txtCGC.SetFocus
        wPessoa = 1
        OptCliente(1).Value = False
    End If

End Sub


Private Sub txtAliqIPI_KeyPress(KeyAscii As Integer)

   VerteclaVirgula txtAliqIPI, KeyAscii

End Sub

Private Sub txtAliqIPI_LostFocus()
   
   If txtAliqIPI.Text <> "" Then
       txtAliqIPI.Text = Format(txtAliqIPI.Text, " ###,###,###,###,#0.00")
   End If

End Sub

Private Sub txtAV_LostFocus()

   If GetAsyncKeyState(vbKeyTab) <> 0 Then

        If Val(cmbCondPagto.Text) > 3 And Val(cmbCondPagto.Text) <> 85 Then
            If txtAV.Text = "" Then
               txtPagtoEnt.SetFocus
               txtAV.Text = 0
            Else
               txtPagtoEnt.SetFocus
            End If
        Else
            txtAV.Text = 0
        End If

   End If
    
End Sub

Private Sub txtBairro_LostFocus()
   txtBairro.Text = UCase(txtBairro.Text)
   
   If txtCFO.Text = 5102 Or txtCFO.Text = 6102 Or txtCFO.Text = 6109 Then
      If Left(cmbCondPagto.Text, 1) <= 3 Then
         If txtCliente.Text = "888888" Then
            If GetAsyncKeyState(vbKeyTab) <> 0 And Trim(txtBairro.Text) = "" Then
               MsgBox "Favor preencher o campo bairro", vbInformation, "Atenção"
               txtBairro.SetFocus
            Else
               Exit Sub
            End If
         End If
      End If
   End If
   
End Sub

Private Sub txtCEP_LostFocus()

   If cmbCondPagto.Text <> "" Then
      If Left(cmbCondPagto.Text, 1) <= 3 Then
         If txtCliente.Text = "888888" Then
            If GetAsyncKeyState(vbKeyTab) <> 0 And Trim(txtCEP.Text) = "" Then
               MsgBox "Favor preencher o campo CEP", vbInformation, "Atenção"
               txtCEP.SetFocus
            Else
               Exit Sub
            End If
         End If
      End If
   End If
   
End Sub

Private Sub txtCFO_LostFocus()

   Dim rdoCodigoOper As rdoResultset
   Dim CondPagto As Long
   Dim Descricao As String
   Dim VendaCompra As String
   
   If GetAsyncKeyState(vbKeyTab) <> 0 Then
       
       If txtCFO.Text <> "" Then
          Set rdoCodigoOper = Conexao.OpenResultset("Select * from CodigoOperacaoNovo where CN_CodigoOperacaoNovo=" & txtCFO.Text & "")
          If rdoCodigoOper.EOF Then
             rdoCodigoOper.Close
             MsgBox "Código de Operação inválido", vbInformation, "Informação"
             txtCFO.SelStart = 0
             txtCFO.SelLength = Len(txtCFO.Text)
             txtCFO.SetFocus
             Exit Sub
             
          End If
          CodigoOperAntigo = rdoCodigoOper("CN_CodigoOperacaoAntigo")
          CodigoOperNovo = rdoCodigoOper("CN_CodigoOperacaoNovo")
          VerificaCodigo
          MontaComboNatureza CodigoOperNovo, cmbNaturezaOperacao
       End If
       
       If GetAsyncKeyState(vbKeyTab) <> 0 And txtCFO.Text = "" Then
          MsgBox "Favor preencher o campo Código de Operação", vbInformation, "Atenção"
          txtCFO.SetFocus
       End If
    '**********************************
       cmbCondPagto.Clear
       If Val(txtCFO.Text) < 500 Then
          VdCp = "C"
       Else
          VdCp = "V"
       End If
       
       Set rdoCombo = Conexao.OpenResultset("Select CP_CodigoCondicao, CP_Descricao, CP_VendaCompra from CondicaoPagto where CP_VendaCompra ='" & VdCp & "' order by CP_CodigoCondicao", Options:=rdExecDirect)
       
       Do While Not rdoCombo.EOF
          CondPagto = rdoCombo("CP_CodigoCondicao")
          Descricao = rdoCombo("CP_Descricao")
          
          If rdoCombo("CP_VendaCompra") = "C" Then
             VendaCompra = "CP"
          Else
             VendaCompra = "VD"
          End If
          
          cmbCondPagto.AddItem CondPagto & " - " & VendaCompra & " - " & Descricao
          cmbCondPagto.ItemData(cmbCondPagto.NewIndex) = CondPagto
          rdoCombo.MoveNext
       Loop
       rdoCombo.Close
    End If
End Sub

Private Sub txtCGC_LostFocus()

   Dim rdoAchaCGCCli As rdoResultset
   Dim rdoAchaCGCLoja
   
   If GetAsyncKeyState(vbKeyTab) <> 0 Then
     If frmCliente.Visible = False Then
        If Left(cmbCondPagto.Text, 1) <> "" Then
           If Left(cmbCondPagto.Text, 1) <= 3 Then
              If txtCliente.Text = "888888" Then
                 If GetAsyncKeyState(vbKeyTab) <> 0 And Trim(txtCGC.Text) = "" Then
                    MsgBox "Favor preencher o campo CGC", vbInformation, "Atenção"
                    txtCGC.SetFocus
                 Else
                    txtInscEst.Enabled = True
                    txtRazao.Enabled = True
                    txtEndereco.Enabled = True
                    txtBairro.Enabled = True
                    txtMunicipio.Enabled = True
                    txtFoneCli.Enabled = True
                    cmbUFCli.Enabled = True
                    txtCEP.Enabled = True
                    txtInscEst.SetFocus
                    'Exit Sub
                 End If
              End If
           End If
        End If
        If txtCFO.Text = 5152 Then
           Set rdoAchaCGCLoja = Conexao.OpenResultset("Select * from Loja where LO_Loja = '" & txtLjDestino.Text & "' and LO_CGC= '" & txtCGC.Text & "' ")
           
           If rdoAchaCGCLoja.EOF Then
              MsgBox "CGC não corresponde à esta loja", vbInformation, "Atenção"
              txtCGC.SelStart = 0
              txtCGC.SelLength = Len(txtCGC.Text)
              txtCGC.SetFocus
              Exit Sub
           Else
              txtInscEst.Text = rdoAchaCGCLoja("LO_InscricaoEstadual")
              txtEndereco.Text = rdoAchaCGCLoja("LO_Endereco")
              txtBairro.Text = rdoAchaCGCLoja("LO_Bairro")
              txtMunicipio.Text = rdoAchaCGCLoja("LO_Municipio")
              txtCEP.Text = rdoAchaCGCLoja("LO_Cep")
              cmbUFCli.Text = rdoAchaCGCLoja("LO_UF")
              If Wserie = "CT" Then
                  txtRazao.Text = "DM MOTORES FERRAMENTAS LTDA"
              Else
                  txtRazao.Text = "DE MEO COMERCIAL IMPORTADORA LTDA"
              End If
           End If
        ElseIf txtCFO.Text = 5202 Or txtCFO.Text = 6202 Then
           Set rdoAchaCGCLoja = Conexao.OpenResultset("Select * from Fornecedor where FO_CGC = '" & txtCGC.Text & "' ")
           
           If Not rdoAchaCGCLoja.EOF Then
              txtInscEst.Text = rdoAchaCGCLoja("FO_InscricaoEstadual")
              txtEndereco.Text = rdoAchaCGCLoja("FO_Endereco")
              txtBairro.Text = "."
              txtMunicipio.Text = rdoAchaCGCLoja("FO_Municipio")
              txtCEP.Text = rdoAchaCGCLoja("FO_Cep")
              cmbUFCli.Text = rdoAchaCGCLoja("FO_Estado")
              txtRazao.Text = rdoAchaCGCLoja("FO_RazaoSocial")
           End If
        Else
           If ClienteExiste = True Then
              If txtCGC.Text <> "" Then
                 Screen.MousePointer = 11
                 
                 Set rdoAchaCGCCli = Conexao.OpenResultset("Select CE_CGC, CE_InscricaoEstadual, " _
                               & "CE_Razao, CE_Endereco, CE_Bairro, CE_Municipio, CE_Estado, " _
                               & "CE_Cep from Cliente where CE_CodigoCliente = " & txtCliente.Text & " AND " _
                               & "CE_CGC= '" & txtCGC.Text & "'")
                               
                 If Not rdoAchaCGCCli.EOF Then
                    txtInscEst.Text = rdoAchaCGCCli("CE_InscricaoEstadual")
                    txtRazao.Text = rdoAchaCGCCli("CE_Razao")
                    txtEndereco.Text = rdoAchaCGCCli("CE_Endereco")
                    txtBairro.Text = rdoAchaCGCCli("CE_Bairro")
                    txtMunicipio.Text = rdoAchaCGCCli("CE_Municipio")
                    'txtUFCli.Text = rdoAchaCGCCli("CE_Estado")
                    txtCEP.Text = rdoAchaCGCCli("CE_Cep")
                    
                    txtreferencia.SetFocus
                 Else
                    MsgBox "CGC incorreto", vbInformation, "Atenção"
                    txtCGC.SelStart = 0
                    txtCGC.SelLength = Len(txtCGC.Text)
                    txtCGC.SetFocus
                 End If
                 rdoAchaCGCCli.Close
                 Screen.MousePointer = 0
              End If
           End If
        End If
      ElseIf frmCliente.Visible = True Then
        MsgBox "Por favor, selecionar se o cliente é pessoa Física ou Jurídica", vbExclamation, "Atenção"
        txtCGC.SetFocus
      Else
        txtInscEst.Enabled = True
        txtInscEst.SetFocus
      End If
   End If
   Screen.MousePointer = 0
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
   Dim rsAchaCliente As rdoResultset
   
   Select Case KeyAscii
   Case 13
        If txtCFO.Text <> "" Then
            If txtCFO.Text = 5102 Or txtCFO.Text = 6102 Then
               If txtCliente.Text = "" Then
                  MsgBox "Favor preencher o campo Cliente", vbInformation, "Atenção"
                  txtCliente.SetFocus
                  Exit Sub
               ElseIf txtCliente.Text = 0 Then
                  MsgBox "Código de cliente inválido", vbInformation, "Atenção"
                  txtCliente.SelStart = 0
                  txtCliente.SelLength = Len(txtCliente.Text)
                  txtCliente.SetFocus
                  Exit Sub
               ElseIf txtCliente.Text = "999999" Then
                  If Left(cmbCondPagto.Text, 1) <= 3 Then
                     txtCGC.Text = "999999999999999"
                     txtInscEst.Text = "999999999999"
                     txtRazao.Text = "VENDA AO CONSUMIDOR"
                     txtEndereco.Text = "."
                     txtBairro.Text = "."
                     txtMunicipio.Text = "SAO PAULO"
                     txtFoneCli.Text = "99999999"
                     cmbUFCli.Text = "SP"
                     txtCEP.Enabled = False
                     txtCGC.Enabled = False
                     txtInscEst.Enabled = False
                     txtRazao.Enabled = False
                     txtEndereco.Enabled = False
                     txtBairro.Enabled = False
                     txtMunicipio.Enabled = False
                     txtFoneCli.Enabled = False
                     cmbUFCli.Enabled = False
                     txtCEP.Enabled = False
                     wPessoa = 2
                     txtreferencia.SetFocus
                  Else
                     MsgBox "Informe um código de cliente para venda faturada", vbInformation, "Informação"
                     txtCliente.SelStart = 0
                     txtCliente.SelLength = Len(txtCliente.Text)
                     txtCliente.SetFocus
                  End If
               ElseIf txtCliente.Text = "888888" Then
                  If Left(cmbCondPagto.Text, 1) <= 3 Then
                     frmCliente.Visible = True
                     txtCEP.Enabled = False
                     txtCGC.Enabled = False
                     txtInscEst.Enabled = False
                     txtRazao.Enabled = False
                     txtEndereco.Enabled = False
                     txtBairro.Enabled = False
                     txtMunicipio.Enabled = False
                     txtFoneCli.Enabled = False
                     cmbUFCli.Enabled = False
                     txtCEP.Enabled = False
                  Else
                     MsgBox "Informe um código de cliente para venda faturada", vbInformation, "Informação"
                     txtCliente.SelStart = 0
                     txtCliente.SelLength = Len(txtCliente.Text)
                     txtCliente.SetFocus
                  End If
               ElseIf txtCliente.Text <> 0 Then
                  Set rsAchaCliente = rdoCnLojaBach.OpenResultset("Select * from Cliente where CE_CodigoCliente = " & txtCliente.Text & "")
               
                  If Not rsAchaCliente.EOF Then
                     If rsAchaCliente("CE_TipoPessoa") = "J" Then
                         wPessoa = 1
                     Else
                         wPessoa = 2
                     End If
                     txtCGC.Text = rsAchaCliente("CE_CGC")
                     txtInscEst.Text = rsAchaCliente("CE_InscricaoEstadual")
                     txtRazao.Text = rsAchaCliente("CE_Razao")
                     txtEndereco.Text = rsAchaCliente("CE_Endereco")
                     txtBairro.Text = rsAchaCliente("CE_Bairro")
                     txtMunicipio.Text = rsAchaCliente("CE_Municipio")
                     txtFoneCli.Text = rsAchaCliente("CE_Telefone")
                     cmbUFCli.Text = rsAchaCliente("CE_Estado")
                     txtCEP.Text = rsAchaCliente("CE_Cep")
                     txtCEP.Enabled = False
                     txtCGC.Enabled = False
                     txtInscEst.Enabled = False
                     txtRazao.Enabled = False
                     txtEndereco.Enabled = False
                     txtBairro.Enabled = False
                     txtMunicipio.Enabled = False
                     txtFoneCli.Enabled = False
                     cmbUFCli.Enabled = False
                     txtCEP.Enabled = False
                     txtreferencia.SetFocus
                     ClienteExiste = True
                  End If
                  rsAchaCliente.Close
               End If
            Else
               txtCliente = 0
            End If
        Else
            MsgBox "Favor colocar o número do CFO", vbInformation, "Atenção"
            txtCFO.SelStart = 0
            txtCFO.SelLength = Len(txtCFO.Text)
            txtCFO.SetFocus
        End If
    Case 9
        MsgBox "Para confirmar o Cliente tecle ENTER", vbInformation, "Atenção"
        txtCliente.SetFocus
    End Select

End Sub

Private Sub txtCliente_LostFocus()
   
   If GetAsyncKeyState(vbKeyTab) <> 0 Then
        MsgBox "Para confirmar o Cliente tecle ENTER", vbInformation, "Atenção"
        txtCliente.SetFocus
   End If

End Sub

Private Sub txtDesconto_GotFocus()
   
   txtDesconto.SelStart = 0
   txtDesconto.SelLength = Len(txtDesconto.Text)

End Sub

Private Sub txtDesconto_KeyPress(KeyAscii As Integer)
   VerteclaVirgula txtDesconto, KeyAscii
End Sub

Private Sub txtDesconto_LostFocus()
   txtDesconto.Text = Format(txtDesconto.Text, " ###,###,###,###,#0.00")
   If Trim(txtDesconto.Text) = "" Then
      txtDesconto.Text = "0,00"
      Desconto = txtDesconto.Text
   Else
      Desconto = txtDesconto.Text
   End If
    If txtDesconto.Text <> " " And txtTotal.Text <> "" Then
        If txtDesconto.Text <> "" And txtTotal.Text <> "" Then
            txtTotal.Text = Format(CDbl(txtTotal.Text) - CDbl(txtDesconto.Text), " ###,###,###,###,#0.00")
        Else
            txtTotal.Text = Format(CDbl(txtTotal.Text), " ###,###,###,###,#0.00")
        End If
    End If

End Sub

Private Sub txtEndereco_LostFocus()
   txtEndereco.Text = UCase(txtEndereco.Text)
   If txtCFO.Text = 5102 Or txtCFO.Text = 6102 Or txtCFO.Text = 5109 Then
      If Left(cmbCondPagto.Text, 1) <= 3 Then
         If txtCliente.Text = "888888" Then
            If GetAsyncKeyState(vbKeyTab) <> 0 And Trim(txtEndereco.Text) = "" Then
               MsgBox "Favor preencher o campo endereço", vbInformation, "Atenção"
               txtEndereco.SetFocus
            Else
               Exit Sub
            End If
         Else
            txtBairro.Enabled = True
            txtBairro.SetFocus
         End If
      End If
   End If
   
End Sub

Private Sub txtFoneCli_LostFocus()
    If GetAsyncKeyState(vbKeyTab) <> 0 Then
        If txtFoneCli.Text = "" Then
            txtFoneCli.Text = 0
            cmbUFCli.SetFocus
        Else
            cmbUFCli.SetFocus
        End If
    End If
End Sub

Private Sub txtFrete_KeyPress(KeyAscii As Integer)
   VerteclaVirgula txtFrete, KeyAscii
End Sub

Private Sub txtFrete_LostFocus()
   txtFrete.Text = Format(txtFrete.Text, " ###,###,###,###,#0.00")
End Sub

Private Sub txtInscEst_LostFocus()
   txtInscEst.Text = UCase(txtInscEst.Text)
   If Left(cmbCondPagto.Text, 1) <> "" Then
      If Left(cmbCondPagto.Text, 1) <= 3 Then
         If txtCliente.Text = "888888" Then
            If GetAsyncKeyState(vbKeyTab) <> 0 And Trim(txtInscEst.Text) = "" Then
               MsgBox "Favor preencher o campo Inscrição Estadual", vbInformation, "Atenção"
               txtInscEst.SetFocus
            Else
               Exit Sub
            End If
         Else
            txtRazao.Enabled = True
            txtRazao.SetFocus
         End If
      End If
   End If
End Sub

Private Sub txtLjDestino_LostFocus()
   Dim rdoLojaDest As rdoResultset
   
   If Trim(txtLjOrigem.Text) <> Trim(txtLjDestino.Text) Then
       If GetAsyncKeyState(vbKeyTab) <> 0 And txtreferencia.Text <> "" Then
          Set rdoLojaDest = Conexao.OpenResultset("Select LO_Loja from Loja where LO_Loja = '" & txtLjDestino.Text & "'")
           
          If rdoLojaDest.EOF Then
             MsgBox "Loja não cadastrada", vbInformation, "Atenção"
             txtLjDestino.SelStart = 0
             txtLjDestino.SelLength = Len(txtLjDestino.Text)
             txtLjDestino.SetFocus
          End If
       ElseIf Trim(txtLjDestino.Text) = "" Then
          MsgBox "Favor digitar o número da loja de destino", , "Atenção"
          txtLjDestino.SelStart = 0
          txtLjDestino.SelLength = Len(txtLjDestino.Text)
          txtLjDestino.SetFocus
       End If
   ElseIf Trim(cmbNaturezaOperacao.Text) = "Transferência" And Trim(txtLjOrigem.Text) = Trim(txtLjDestino.Text) Then
       MsgBox "Loja destino não pode ser igual loja origem.", , "Atenção"
       txtLjDestino.SelStart = 0
       txtLjDestino.SelLength = Len(txtLjDestino.Text)
       txtLjDestino.SetFocus
   Else
       txtOutroVendedor.Text = 0
       txtOutroVendedor.Enabled = False
   End If
   
End Sub

Private Sub txtLjOrigem_LostFocus()
   Dim rdoProcura As rdoResultset
   Dim rdoLojaOri As rdoResultset
   
   If GetAsyncKeyState(vbKeyTab) <> 0 And txtLjOrigem.Text <> "" Then
      Screen.MousePointer = 11
      
      Set rdoLojaOri = Conexao.OpenResultset("Select LO_Loja from Loja where LO_Loja = '" & txtLjOrigem.Text & "'")
      
      If rdoLojaOri.EOF Then
         Screen.MousePointer = 0
         rdoLojaOri.Close
         MsgBox "Loja não cadastrada", vbInformation, "Atenção"
         txtLjOrigem.SelStart = 0
         txtLjOrigem.SelLength = Len(txtLjOrigem.Text)
         txtLjOrigem.SetFocus
         Exit Sub
      End If
      rdoLojaOri.Close
      
      CargaNota
      
      Screen.MousePointer = 0
   End If
   
End Sub

Private Sub txtMunicipio_LostFocus()
   txtMunicipio.Text = UCase(txtMunicipio.Text)
   If txtCFO.Text = 5102 Or txtCFO.Text = 6102 Or txtCFO.Text = 6109 Then
      If Left(cmbCondPagto.Text, 1) <= 3 Then
         If txtCliente.Text = "888888" Then
            If GetAsyncKeyState(vbKeyTab) <> 0 And Trim(txtMunicipio.Text) = "" Then
               MsgBox "Favor preencher o campo Município", vbInformation, "Atenção"
               txtMunicipio.SetFocus
            Else
               txtFoneCli.SetFocus
               Exit Sub
            End If
         End If
      End If
   End If
   
End Sub

Private Sub txtNF_LostFocus()
'   If txtNF <> "" Then
'      Set RdoNota = Conexao.openresultset("Select capanfvenda.*, itemnfvenda.*, PR_Descricao from CapaNFVenda, ItemnfVenda, Produto where PR_Referencia=VI_Referencia and VC_NotaFiscal=VI_NotaFiscal and VC_Serie= VI_Serie and VC_NotaFiscal=" & txtNF.Text & " and VC_Serie= '" & txtSerie.Text & "' and VC_LojaOrigem = '" & txtLjOrigem.Text & "' and VC_CodigoOperacao=" & txtCFO.Text & " ")
'
'      If Not RdoNota.EOF Then
'         AlteraDados = True
'
'         VerificaCodigo
'
'         If txtCFO.Text = 132 Or txtCFO.Text = 232 Then
'            txtLjOrigem.Text = RdoNota("VC_LojaOrigem")
'            txtVendedor.Text = RdoNota("VC_CodigoVendedor")
'            mskDataEmi.Text = Format(RdoNota("VC_DataEmissao"), "dd/mm/yyyy")
'            txtCliente.Text = RdoNota("VC_Cliente")
'            txtCGC.Text = IIf(IsNull(RdoNota("VC_CGCCliente")), 0, RdoNota("VC_CGCCliente"))
'            txtInscEst.Text = IIf(IsNull(RdoNota("VC_InscEstCliente")), 0, RdoNota("VC_InscEstCliente"))
'            txtRazao.Text = IIf(IsNull(RdoNota("VC_NomeCliente")), 0, RdoNota("VC_NomeCliente"))
'            txtEndereco.Text = IIf(IsNull(RdoNota("VC_EnderecoCliente")), "", RdoNota("VC_EnderecoCliente"))
'            txtBairro.Text = IIf(IsNull(RdoNota("VC_BairroCliente")), "", RdoNota("VC_BairroCliente"))
'            txtMunicipio.Text = IIf(IsNull(RdoNota("VC_MunicipioCliente")), "", RdoNota("VC_MunicipioCliente"))
'            cmbUFCli.Text = IIf(IsNull(RdoNota("VC_UFCliente")), "SP", RdoNota("VC_UFCliente"))
'            txtCEP.Text = RdoNota("VC_CEPCliente")
'            txtSubTotal.Text = Format(RdoNota("VC_ValorMercadorias"), "###,###,###,#0.00")
'            txtBaseICMS.Text = Format(RdoNota("VC_BaseICMS"), "###,###,###,#0.00")
'            txtAliqICMS.Text = Format(RdoNota("VC_AliquotaICMS"), "###,###,###,#0.00")
'            txtValorICMS.Text = Format(RdoNota("VC_ValorICMS"), "###,###,###,#0.00")
'            txtAliqIPI.Text = Format(RdoNota("VI_AliquotaIPI"), "###,###,###,#0.00")
'            txtValorIPI.Text = Format(RdoNota("VC_ValorIPI"), "###,###,###,#0.00")
'            txtTotal.Text = Format(RdoNota("VC_TotalNota"), "###,###,###,#0.00")
'
'            Do While Not RdoNota.EOF
'               grdItens.AddItem RdoNota("VI_Referencia") & Chr(9) & RdoNota("PR_Descricao") & Chr(9) & RdoNota("VI_Quantidade") & Chr(9) & RdoNota("VI_PrecoUnitario") & Chr(9) & RdoNota("VI_ValorMercadoria")
'               RdoNota.MoveNext
'            Loop
'
'            If grdItens.Rows > 2 And grdItens.TextMatrix(1, 0) = "" Then
'               grdItens.RemoveItem 1
'            End If
'
'            'txtLjOri.SetFocus
'
'         End If
'      End If
'      RdoNota.Close
'   End If
    If txtNF.Text <> "" Then
        txtSerie.SetFocus
    End If

End Sub

Private Sub txtOutroVendedor_LostFocus()
   
   If GetAsyncKeyState(vbKeyTab) <> 0 Then
      If txtOutroVendedor.Text <> "" And txtLjOrigem.Text <> txtLjDestino.Text Then
         Set rdoVendedor = Conexao.OpenResultset("Select VE_CodigoVendedor, VE_Nome from Vendedor where VE_CodigoVendedor= " & txtOutroVendedor.Text & " and VE_Loja = '" & txtLjDestino.Text & "'")
         
         If rdoVendedor.EOF Then
            MsgBox "Vendedor não encontrado.", vbInformation, "Atenção"
            txtOutroVendedor.SelStart = 0
            txtOutroVendedor.SelLength = Len(txtVendedor.Text)
            txtOutroVendedor.SetFocus
         Else
            txtOutroVendedor.Text = rdoVendedor("VE_CodigoVendedor") & " - " & rdoVendedor("VE_Nome")
         End If
         
         rdoVendedor.Close
      End If
   End If

End Sub

Private Sub txtPagtoEnt_KeyPress(KeyAscii As Integer)
   VerteclaVirgula txtPagtoEnt, KeyAscii
End Sub

Private Sub txtPagtoEnt_LostFocus()
   If GetAsyncKeyState(vbKeyTab) <> 0 Then
   
      If txtPagtoEnt.Text <> "" Or txtPagtoEnt.Text <> "0.00" Then
         If Val(cmbCondPagto) = 3 Then
            txtPagtoEnt.Text = Format(txtPagtoEnt.Text, " ###,###,###,###,#0.00")
         ElseIf Val(cmbCondPagto) > 3 Then
            txtPagtoEnt.Text = Format(txtPagtoEnt.Text, " ###,###,###,###,#0.00")
         Else
            txtPagtoEnt.Text = "0.00"
            txtPagtoEnt.Text = Format(txtPagtoEnt.Text, " ###,###,###,###,#0.00")
         End If
      Else
         txtPagtoEnt.Text = "0.00"
         txtPagtoEnt.Text = Format(txtPagtoEnt.Text, " ###,###,###,###,#0.00")
      End If
      txtCliente.SetFocus
   
   End If
End Sub

Private Sub txtPrecoUnit_Change()
   CalculaTotal
End Sub

Private Sub txtPrecoUnit_LostFocus()
    If GetAsyncKeyState(vbKeyTab) <> 0 And Not IsNull(txtPrecoUnit) And CDbl(txtPrecoUnit) <> 0 Then
       PreencheGrid
       
       
       SubTotal = Format(SubTotal + CDbl(txtTotalItem.Text), " ###,###,###,###,#0.00")
       
       txtSubTotal.Text = Format(SubTotal, " ###,###,###,###,#0.00")
       
       txtTotal.Text = Format(SubTotal, " ###,###,###,###,#0.00")
       
       txtreferencia.Text = ""
       txtdescricao.Text = ""
       txtQuant.Text = ""
       txtPrecoUnit.Text = ""
       txtTotalItem.Text = ""
       txtQuant.Enabled = False
       txtPrecoUnit.Enabled = False
       txtreferencia.SetFocus
    Else
        If GetAsyncKeyState(vbKeyTab) = 9 Then
            MsgBox "Favor colocar o preço unitário", vbInformation, "Atenção"
            txtPrecoUnit.SetFocus
        End If
    End If
End Sub

Private Sub txtQuant_Change()
   CalculaTotal
End Sub

Private Sub txtQuant_KeyPress(KeyAscii As Integer)
   vertecla KeyAscii
End Sub

Private Sub txtQuant_LostFocus()

    If GetAsyncKeyState(vbKeyTab) <> 0 Then
        If Not IsNull(txtQuant) And Val(txtQuant) <> 0 Then
            txtPrecoUnit.Enabled = True
            txtPrecoUnit.SetFocus
        ElseIf txtreferencia.Text <> "" Then
            MsgBox "Favor digitar a quantidade", vbCritical, "Atenção"
            txtQuant.SetFocus
        Else
            txtDesconto.SetFocus
        End If
    End If

End Sub

Private Sub txtRazao_LostFocus()
   txtRazao.Text = UCase(txtRazao.Text)
   If txtCFO.Text = 5102 Or txtCFO.Text = 6102 Or txtCFO.Text = 5109 Then
      If Left(cmbCondPagto.Text, 1) <= 3 Then
         If txtCliente.Text = "888888" Then
            If GetAsyncKeyState(vbKeyTab) <> 0 And Trim(txtRazao.Text) = "" Then
               MsgBox "Favor preencher o campo Razão", vbInformation, "Atenção"
               txtRazao.SetFocus
            Else
               Exit Sub
            End If
         Else
            txtEndereco.Enabled = True
            txtEndereco.SetFocus
         End If
      End If
   End If
End Sub

Private Sub txtReferencia_KeyPress(KeyAscii As Integer)
   If txtreferencia.Text = "" And KeyAscii = vbKeyReturn Then
      If txtCFO.Text <> 5152 And txtCFO.Text <> 5910 Then
         KeyAscii = 0
         txtDesconto.SetFocus
      Else
         KeyAscii = 0
         txtTotal.SetFocus
      End If
   End If
End Sub

Private Sub txtReferencia_LostFocus()
   Dim rdoAchaReferencia As rdoResultset
   
    If Altera = 0 Then
        If txtreferencia.Text <> "" Then
            For x = 1 To grdItens.Rows - 1
               If grdItens.TextMatrix(x, 0) = txtreferencia.Text Then
                  MsgBox "Referência já digitada.", vbCritical, "Atenção"
                  txtreferencia.SelStart = 0
                  txtreferencia.SelLength = Len(txtreferencia.Text)
                  txtreferencia.SetFocus
                  Exit Sub
               End If
            Next x
               Set rdoAchaReferencia = Conexao.OpenResultset("Select PR_Descricao from Produto where PR_Referencia= '" & txtreferencia.Text & "'")
         
               If Not rdoAchaReferencia.EOF Then
                  txtdescricao.Text = rdoAchaReferencia("PR_Descricao")
                  txtQuant.Enabled = True
                  txtQuant.SetFocus
               Else
                  MsgBox "Referência não cadastrada", vbInformation, "Atenção"
                  txtreferencia.SelStart = 0
                  txtreferencia.SelLength = Len(txtreferencia.Text)
                  txtreferencia.SetFocus
               End If
        ElseIf grdItens.Rows = 2 And grdItens.TextMatrix(1, 0) = "" And txtCliente.Text <> "" Then
           MsgBox "Favor incluir um item", vbCritical, "Atenção"
           txtreferencia.SetFocus
        Else
           txtDesconto.Enabled = True
           txtDesconto.SelStart = 0
           txtDesconto.SelLength = Len(txtDesconto)
           txtDesconto.SetFocus
        End If
    End If
End Sub

Sub PreencheGrid()
   Dim rdoPeso As rdoResultset
   
   If Altera = 1 Then
      grdItens.TextMatrix(grdItens.Row, 0) = txtreferencia.Text
      grdItens.TextMatrix(grdItens.Row, 1) = txtdescricao.Text
      grdItens.TextMatrix(grdItens.Row, 2) = txtQuant.Text
      grdItens.TextMatrix(grdItens.Row, 3) = Format(txtPrecoUnit.Text, " ###,###,###,###,#0.00")
      grdItens.TextMatrix(grdItens.Row, 4) = Format(txtTotalItem.Text, " ###,###,###,###,#0.00")
   Else
      grdItens.AddItem txtreferencia.Text & Chr(vbKeyTab) & txtdescricao.Text & Chr(vbKeyTab) & txtQuant.Text & Chr(vbKeyTab) & Format(txtPrecoUnit.Text, " ###,###,###,###,#0.00") & Chr(vbKeyTab) & Format(txtTotalItem.Text, " ###,###,###,###,#0.00")
      
      If grdItens.Rows > 2 And grdItens.TextMatrix(1, 0) = "" Then
         grdItens.RemoveItem 1
      End If
      If grdItens.Rows > 6 And grdItens.TextMatrix(1, 0) <> "" Then
         grdItens.TopRow = grdItens.Rows - 1
      End If
      
      cmdGrava.Enabled = True
      
   End If
   
   Altera = 0
End Sub

Private Sub txtSerie_GotFocus()

    If txtCliente.Text = "888888" Then
        txtCGC.SetFocus
    End If

End Sub

Private Sub txtSerie_LostFocus()

    txtSerie.Text = UCase(txtSerie.Text)

End Sub

Private Sub CargaNota()

   On Error Resume Next
   
   Dim Nf As String
   
   If Val(txtNF.Text) > 0 And Len(Trim(txtSerie.Text)) = 2 And Trim(txtLjOrigem.Text) <> "" Then
       Screen.MousePointer = 11
       
                       
               If txtSerie.Text = "CF" Then
                  Nf = txtNF.Text & glb_ECF
               Else
                  Nf = txtNF.Text
               End If
                       
               Set RdoNota = Conexao.OpenResultset("Select VC_NotaFiscal, VC_Serie, VC_LojaOrigem, VC_LojaDestino, VC_CGCLojaDestino, " _
                           & "VC_DataEmissao, VC_HoraEmissao, VC_TipoNota, VC_Cliente, VC_CodigoOperacao, VC_CondicaoPagamento, VC_AV, " _
                           & "VC_DataPagamento, VC_PagamentoEntrada, VC_CodigoVendedor, VC_NumeroPedido, VC_TotalNota, VC_Desconto, " _
                           & "VC_ValorMercadorias, VC_BaseICMS, VC_AliquotaICMS, VC_ValorICMS, VC_ValorFrete, VC_ValorFreteCobrado, " _
                           & "VC_ValorIPI, VC_NomeCliente, VC_EnderecoCliente, VC_BairroCliente, VC_MunicipioCliente, VC_UFCliente, " _
                           & "VC_CEPCliente, VC_TelefoneCliente, VC_EndEntregaCliente, VC_CGCCliente, VC_InscEstCliente, VC_PedidoCliente, " _
                           & "VC_Dinheiro, VC_Cheque, VC_ChequePre, VC_Cartao, VC_Deposito, VC_NotaCredito, VC_Financiada, VC_Faturada, " _
                           & "VC_AVistaReceber, VC_PesoBruto, VC_PesoLiquido, VC_Observacao, VC_SituacaoComunicacao, VC_Situacao, " _
                           & "VC_EncargosFinanceiros, ITEMNFVENDA.*, PR_DESCRICAO from CapaNFVenda, ITEMNFVENDA,PRODUTO where " _
                           & "PR_REFERENCIA=VI_REFERENCIA AND VC_NOTAFISCAL=VI_NOTAFISCAL AND VC_SERIE= VI_SERIE AND VC_LOJAORIGEM=VI_LOJAORIGEM AND " _
                           & "VC_NotaFiscal=" & Val(Nf) & " and VC_Serie='" & Trim(txtSerie.Text) & "' and VC_LojaOrigem = '" & Trim(txtLjOrigem.Text) & "'")
                  
                  If Not RdoNota.EOF Then
                  
                      AlteraDados = True
                      
                      txtCFO.Text = RdoNota("VC_CodigoOperacao")
                      
                      MontaComboNatureza txtCFO.Text, cmbNaturezaOperacao
                      
                      cmbNaturezaOperacao.ListIndex = 0
                      
                      VerificaCodigo
                
                      Select Case txtCFO.Text
                         Case 512
                            txtLjOrigem.Text = Trim(RdoNota("VC_LojaOrigem"))
                            txtLjDestino.Text = Trim(RdoNota("VC_LojaDestino"))
                            mskDataEmi.Text = Format(RdoNota("VC_DataEmissao"), "dd/mm/yyyy")
                            txtCGC.Text = RdoNota("VC_CGCLojaDestino")
                            txtInscEst.Text = IIf(IsNull(RdoNota("VC_InscEstCliente")), 0, RdoNota("VC_InscEstCliente"))
                            txtSubTotal.Text = Format(RdoNota("VC_ValorMercadorias"), "###,###,###,##0.00")
                            txtAliqIPI.Text = Format(RdoNota("VC_AliquotaIPI"), "###,###,###,##0.00")
                            txtValorIPI.Text = Format(RdoNota("VC_ValorIPI"), "###,###,###,##0.00")
                            txtTotal.Text = Format(RdoNota("VC_TotalNota"), "###,###,###,#0.00")
                   
                         Case 512, 612
                            txtLjOrigem.Text = RdoNota("VC_LojaOrigem")
                            txtVendedor.Text = RdoNota("VC_CodigoVendedor")
                            mskDataEmi.Text = Format(RdoNota("VC_DataEmissao"), "dd/mm/yyyy")
                            cmbCondPagto.Text = RdoNota("VC_CondicaoPagamento")
                            txtPagtoEnt.Text = IIf(IsNull(RdoNota("VC_PagamentoEntrada")), 0, RdoNota("VC_PagamentoEntrada"))
                            txtCliente.Text = RdoNota("VC_Cliente")
                            txtCGC.Text = IIf(IsNull(RdoNota("VC_CGCCliente")), 0, RdoNota("VC_CGCCliente"))
                            txtInscEst.Text = IIf(IsNull(RdoNota("VC_InscEstCliente")), 0, RdoNota("VC_InscEstCliente"))
                            txtRazao.Text = IIf(IsNull(RdoNota("VC_NomeCliente")), 0, RdoNota("VC_NomeCliente"))
                            txtEndereco.Text = IIf(IsNull(RdoNota("VC_EnderecoCliente")), "", RdoNota("VC_EnderecoCliente"))
                            txtBairro.Text = IIf(IsNull(RdoNota("VC_BairroCliente")), "", RdoNota("VC_BairroCliente"))
                            txtMunicipio.Text = IIf(IsNull(RdoNota("VC_MunicipioCliente")), "", RdoNota("VC_MunicipioCliente"))
                            cmbUFCli.Text = IIf(IsNull(RdoNota("VC_UFCliente")), "SP", RdoNota("VC_UFCliente"))
                            txtCEP.Text = RdoNota("VC_CEPCliente")
                            txtSubTotal.Text = Format(RdoNota("VC_ValorMercadorias"), "###,###,###,#0.00")
                            txtDesconto.Text = Format(RdoNota("VC_Desconto"), "###,###,###,#0.00")
                            txtAliqIPI.Text = Format(RdoNota("VI_AliquotaIPI"), "###,###,###,#0.00")
                            txtValorIPI.Text = Format(RdoNota("VC_ValorIPI"), "###,###,###,#0.00")
                            txtFrete.Text = Format(RdoNota("VC_ValorFrete"), "###,###,###,#0.00")
                            txtTotal.Text = Format(RdoNota("VC_TotalNota"), "###,###,###,#0.00")
                   
                         Case 599, 699, 199, 299
                            txtLjOrigem.Text = RdoNota("VC_LojaOrigem")
                            mskDataEmi.Text = Format(RdoNota("VC_DataEmissao"), "dd/mm/yyyy")
                            txtCGC.Text = IIf(IsNull(RdoNota("VC_CGCCliente")), 0, RdoNota("VC_CGCCliente"))
                            txtInscEst.Text = IIf(IsNull(RdoNota("VC_InscEstCliente")), 0, RdoNota("VC_InscEstCliente"))
                            txtRazao.Text = IIf(IsNull(RdoNota("VC_NomeCliente")), 0, RdoNota("VC_NomeCliente"))
                            txtEndereco.Text = IIf(IsNull(RdoNota("VC_EnderecoCliente")), "", RdoNota("VC_EnderecoCliente"))
                            txtBairro.Text = IIf(IsNull(RdoNota("VC_BairroCliente")), "", RdoNota("VC_BairroCliente"))
                            txtMunicipio.Text = IIf(IsNull(RdoNota("VC_MunicipioCliente")), "", RdoNota("VC_MunicipioCliente"))
                            cmbUFCli.Text = IIf(IsNull(RdoNota("VC_UFCliente")), "SP", RdoNota("VC_UFCliente"))
                            txtCEP.Text = RdoNota("VC_CEPCliente")
                            txtSubTotal.Text = Format(RdoNota("VC_ValorMercadorias"), "###,###,###,#0.00")
                            txtDesconto.Text = Format(RdoNota("VC_Desconto"), "###,###,###,#0.00")
                            txtAliqIPI.Text = Format(RdoNota("VI_AliquotaIPI"), "###,###,###,#0.00")
                            txtValorIPI.Text = Format(RdoNota("VC_ValorIPI"), "###,###,###,#0.00")
                            txtFrete.Text = Format(RdoNota("VC_ValorFrete"), "###,###,###,#0.00")
                            txtTotal.Text = Format(RdoNota("VC_TotalNota"), "###,###,###,#0.00")
                
                         Case 522, 622
                            txtLjOrigem.Text = RdoNota("VC_LojaOrigem")
                            txtVendedor.Text = RdoNota("VC_CodigoVendedor")
                            mskDataEmi.Text = Format(RdoNota("VC_DataEmissao"), "dd/mm/yyyy")
                            txtCGC.Text = IIf(IsNull(RdoNota("VC_CGCCliente")), 0, RdoNota("VC_CGCCliente"))
                            txtInscEst.Text = IIf(IsNull(RdoNota("VC_InscEstCliente")), 0, RdoNota("VC_InscEstCliente"))
                            txtRazao.Text = IIf(IsNull(RdoNota("VC_NomeCliente")), 0, RdoNota("VC_NomeCliente"))
                            txtEndereco.Text = IIf(IsNull(RdoNota("VC_EnderecoCliente")), "", RdoNota("VC_EnderecoCliente"))
                            txtBairro.Text = IIf(IsNull(RdoNota("VC_BairroCliente")), "", RdoNota("VC_BairroCliente"))
                            txtMunicipio.Text = IIf(IsNull(RdoNota("VC_MunicipioCliente")), "", RdoNota("VC_MunicipioCliente"))
                            cmbUFCli.Text = IIf(IsNull(RdoNota("VC_UFCliente")), "SP", RdoNota("VC_UFCliente"))
                            txtCEP.Text = RdoNota("VC_CEPCliente")
                            txtSubTotal.Text = Format(RdoNota("VC_ValorMercadorias"), "###,###,###,#0.00")
                            txtDesconto.Text = Format(RdoNota("VC_Desconto"), "###,###,###,#0.00")
                            txtAliqIPI.Text = Format(RdoNota("VI_AliquotaIPI"), "###,###,###,#0.00")
                            txtValorIPI.Text = Format(RdoNota("VC_ValorIPI"), "###,###,###,#0.00")
                            txtTotal.Text = Format(RdoNota("VC_TotalNota"), "###,###,###,#0.00")
                         
                         Case 122, 222
                            txtLjOrigem.Text = RdoNota("VC_LojaOrigem")
                            txtVendedor.Text = RdoNota("VC_CodigoVendedor")
                            mskDataEmi.Text = Format(RdoNota("VC_DataEmissao"), "dd/mm/yyyy")
                            txtCGC.Text = IIf(IsNull(RdoNota("VC_CGCCliente")), 0, RdoNota("VC_CGCCliente"))
                            txtInscEst.Text = IIf(IsNull(RdoNota("VC_InscEstCliente")), 0, RdoNota("VC_InscEstCliente"))
                            txtRazao.Text = IIf(IsNull(RdoNota("VC_NomeCliente")), 0, RdoNota("VC_NomeCliente"))
                            txtEndereco.Text = IIf(IsNull(RdoNota("VC_EnderecoCliente")), "", RdoNota("VC_EnderecoCliente"))
                            txtBairro.Text = IIf(IsNull(RdoNota("VC_BairroCliente")), "", RdoNota("VC_BairroCliente"))
                            txtMunicipio.Text = IIf(IsNull(RdoNota("VC_MunicipioCliente")), "", RdoNota("VC_MunicipioCliente"))
                            cmbUFCli.Text = IIf(IsNull(RdoNota("VC_UFCliente")), "SP", RdoNota("VC_UFCliente"))
                            txtCEP.Text = RdoNota("VC_CEPCliente")
                            txtSubTotal.Text = Format(RdoNota("VC_ValorMercadorias"), "###,###,###,#0.00")
                            txtDesconto.Text = Format(RdoNota("VC_Desconto"), "###,###,###,#0.00")
                            txtAliqIPI.Text = Format(RdoNota("VI_AliquotaIPI"), "###,###,###,#0.00")
                            txtValorIPI.Text = Format(RdoNota("VC_ValorIPI"), "###,###,###,#0.00")
                            txtTotal.Text = Format(RdoNota("VC_TotalNota"), "###,###,###,#0.00")
                      End Select
                      
                      Do While Not RdoNota.EOF
                         grdItens.AddItem RdoNota("VI_Referencia") & Chr(9) & RdoNota("PR_Descricao") & Chr(9) & RdoNota("VI_Quantidade") & Chr(9) & Format(RdoNota("VI_PrecoUnitario"), "###,###,###,#0.00") & Chr(9) & Format(RdoNota("VI_ValorMercadoria"), "###,###,###,#0.00")
                         RdoNota.MoveNext
                      Loop
                      
                      If grdItens.Rows > 2 And grdItens.TextMatrix(1, 0) = "" Then
                         grdItens.RemoveItem 1
                      End If
                      
                      RdoNota.Close
                  Else
                      SQL = "Select NFCapa.*, NFItens.*, PR_DESCRICAO from NFCapa, NFItens,PRODUTO where " _
                          & "PR_REFERENCIA=NFItens.REFERENCIA AND NFCapa.NF = NFItens.NF AND NFCapa.SERIE = NFItens.SERIE AND NFCapa.LOJAORIGEM = NFItens.LOJAORIGEM AND " _
                          & "NFCapa.NF=" & Val(txtNF.Text) & " and NFCapa.Serie='" & Trim(txtSerie.Text) & "' and NFCapa.LojaOrigem = '" & Trim(txtLjOrigem.Text) & "'"
                          
                      Set RdoNota = rdoCnLojaBach.OpenResultset(SQL)
                      If Not RdoNota.EOF Then
                          txtCFO.Text = Trim(RdoNota("CFOAux"))
                      
                          AlteraDados = True
                          
                          MontaComboNatureza txtCFO.Text, cmbNaturezaOperacao
                          
                          cmbNaturezaOperacao.ListIndex = 0
                          
                          VerificaCodigo
                    
                          Select Case txtCFO.Text
                             Case 5152
                                txtLjOrigem.Text = Trim(RdoNota("NfCapa.LojaOrigem"))
                                txtLjDestino.Text = Trim(RdoNota("LojaT"))
                                mskDataEmi.Text = Format(RdoNota("NfCapa.DataEmi"), "dd/mm/yyyy")
                                txtCGC.Text = RdoNota("CGCCli")
                                txtInscEst.Text = IIf(IsNull(RdoNota("INSCRICLI")), 0, RdoNota("INSCRICLI"))
                                txtSubTotal.Text = Format(RdoNota("VLRMERCADORIA"), "###,###,###,##0.00")
                                txtAliqIPI.Text = Format(0, "###,###,###,##0.00")
                                txtValorIPI.Text = Format(RdoNota("TOTALIPI"), "###,###,###,##0.00")
                                txtTotal.Text = Format(RdoNota("TotalNota"), "###,###,###,#0.00")
                       
                             Case 5102, 6102
                                txtLjOrigem.Text = RdoNota("NfCapa.LojaOrigem")
                                txtVendedor.Text = RdoNota("NfCapa.VENDEDOR")
                                mskDataEmi.Text = Format(RdoNota("NfCapa.DataEmi"), "dd/mm/yyyy")
                                cmbCondPagto.Text = RdoNota("CONDPAG")
                                txtPagtoEnt.Text = IIf(IsNull(RdoNota("PGENTRA")), 0, RdoNota("PGENTRA"))
                                txtCliente.Text = RdoNota("NFCapa.Cliente")
                                txtCGC.Text = IIf(IsNull(RdoNota("CGCCli")), 0, RdoNota("CGCCli"))
                                txtInscEst.Text = IIf(IsNull(RdoNota("INSCRICLI")), 0, RdoNota("INSCRICLI"))
                                txtRazao.Text = IIf(IsNull(RdoNota("NOMCLI")), 0, RdoNota("NOMCLI"))
                                txtEndereco.Text = IIf(IsNull(RdoNota("ENDCLI")), "", RdoNota("ENDCLI"))
                                txtBairro.Text = IIf(IsNull(RdoNota("BAIRROCLI")), "", RdoNota("BAIRROCLI"))
                                txtMunicipio.Text = IIf(IsNull(RdoNota("MUNICIPIOCLI")), "", RdoNota("MUNICIPIOCLI"))
                                cmbUFCli.Text = IIf(IsNull(RdoNota("UFCLIENTE")), "SP", RdoNota("UFCLIENTE"))
                                txtCEP.Text = RdoNota("CEPCLI")
                                txtSubTotal.Text = Format(RdoNota("VLRMERCADORIA"), "###,###,###,#0.00")
                                txtDesconto.Text = Format(RdoNota("NfCapa.DESCONTO"), "###,###,###,#0.00")
                                txtAliqIPI.Text = Format(0, "###,###,###,#0.00")
                                txtValorIPI.Text = Format(RdoNota("TOTALIPI"), "###,###,###,#0.00")
                                txtFrete.Text = Format(RdoNota("FRETECOBR"), "###,###,###,#0.00")
                                txtTotal.Text = Format(RdoNota("TOTALNOTA"), "###,###,###,#0.00")
                       
                             Case 5910, 6910, 1910, 2910
                                txtLjOrigem.Text = RdoNota("NfCapa.LojaOrigem")
                                mskDataEmi.Text = Format(RdoNota("NfCapa.DataEmi"), "dd/mm/yyyy")
                                txtCGC.Text = IIf(IsNull(RdoNota("CGCCli")), 0, RdoNota("CGCCli"))
                                txtInscEst.Text = IIf(IsNull(RdoNota("INSCRICLI")), 0, RdoNota("INSCRICLI"))
                                txtRazao.Text = IIf(IsNull(RdoNota("NOMCLI")), 0, RdoNota("NOMCLI"))
                                txtEndereco.Text = IIf(IsNull(RdoNota("ENDCLI")), "", RdoNota("ENDCLI"))
                                txtBairro.Text = IIf(IsNull(RdoNota("BAIRROCLI")), "", RdoNota("BAIRROCLI"))
                                txtMunicipio.Text = IIf(IsNull(RdoNota("MUNICIPIOCLI")), "", RdoNota("MUNICIPIOCLI"))
                                cmbUFCli.Text = IIf(IsNull(RdoNota("UFCLIENTE")), "SP", RdoNota("UFCLIENTE"))
                                txtCEP.Text = RdoNota("CEPCLI")
                                txtSubTotal.Text = Format(RdoNota("VLRMERCADORIA"), "###,###,###,#0.00")
                                txtDesconto.Text = Format(RdoNota("NfCapa.Desconto"), "###,###,###,#0.00")
                                txtAliqIPI.Text = Format(0, "###,###,###,#0.00")
                                txtValorIPI.Text = Format(RdoNota("TOTALIPI"), "###,###,###,#0.00")
                                txtFrete.Text = Format(RdoNota("FRETECOBR"), "###,###,###,#0.00")
                                txtTotal.Text = Format(RdoNota("TOTALNOTA"), "###,###,###,#0.00")
                    
                             Case 5202, 6202
                                txtLjOrigem.Text = RdoNota("NfCapa.LojaOrigem")
                                txtVendedor.Text = RdoNota("NfCapa.Vendedor")
                                mskDataEmi.Text = Format(RdoNota("NfCapa.DataEmi"), "dd/mm/yyyy")
                                txtCGC.Text = IIf(IsNull(RdoNota("CGCCli")), 0, RdoNota("CGCCli"))
                                txtInscEst.Text = IIf(IsNull(RdoNota("InscriCli")), 0, RdoNota("InscriCli"))
                                txtRazao.Text = IIf(IsNull(RdoNota("NomCli")), 0, RdoNota("NomCli"))
                                txtEndereco.Text = IIf(IsNull(RdoNota("EndCli")), "", RdoNota("EndCli"))
                                txtBairro.Text = IIf(IsNull(RdoNota("BairroCli")), "", RdoNota("BairroCli"))
                                txtMunicipio.Text = IIf(IsNull(RdoNota("MunicipioCli")), "", RdoNota("MunicipioCli"))
                                cmbUFCli.Text = IIf(IsNull(RdoNota("UFCliente")), "SP", RdoNota("UFCliente"))
                                txtCEP.Text = RdoNota("CEPCli")
                                txtSubTotal.Text = Format(RdoNota("VLRMERCADORIA"), "###,###,###,#0.00")
                                txtDesconto.Text = Format(RdoNota("NfCapa.Desconto"), "###,###,###,#0.00")
                                txtAliqIPI.Text = Format(0, "###,###,###,#0.00")
                                txtValorIPI.Text = Format(RdoNota("TOTALIPI"), "###,###,###,#0.00")
                                txtTotal.Text = Format(RdoNota("TOTALNOTA"), "###,###,###,#0.00")
                             
                             Case 1202, 2202
                                txtLjOrigem.Text = RdoNota("NfCapa.LojaOrigem")
                                txtVendedor.Text = RdoNota("NfCapa.Vendedor")
                                mskDataEmi.Text = Format(RdoNota("NfCapa.DataEmi"), "dd/mm/yyyy")
                                txtCGC.Text = IIf(IsNull(RdoNota("CGCCli")), 0, RdoNota("CGCCli"))
                                txtInscEst.Text = IIf(IsNull(RdoNota("InscriCli")), 0, RdoNota("InscriCli"))
                                txtRazao.Text = IIf(IsNull(RdoNota("NomCli")), 0, RdoNota("NomCli"))
                                txtEndereco.Text = IIf(IsNull(RdoNota("EndCli")), "", RdoNota("EndCli"))
                                txtBairro.Text = IIf(IsNull(RdoNota("BairroCli")), "", RdoNota("BairroCli"))
                                txtMunicipio.Text = IIf(IsNull(RdoNota("MunicipioCli")), "", RdoNota("MunicipioCli"))
                                cmbUFCli.Text = IIf(IsNull(RdoNota("UFCliente")), "SP", RdoNota("UFCliente"))
                                txtCEP.Text = RdoNota("CEPCli")
                                txtSubTotal.Text = Format(RdoNota("VLRMERCADORIA"), "###,###,###,#0.00")
                                txtDesconto.Text = Format(RdoNota("NfCapa.Desconto"), "###,###,###,#0.00")
                                txtAliqIPI.Text = Format(0, "###,###,###,#0.00")
                                txtValorIPI.Text = Format(RdoNota("TOTALIPI"), "###,###,###,#0.00")
                                txtTotal.Text = Format(RdoNota("TOTALNOTA"), "###,###,###,#0.00")
                          End Select
                          
                          Do While Not RdoNota.EOF
                             grdItens.AddItem RdoNota("Referencia") & Chr(9) & RdoNota("PR_Descricao") & Chr(9) & RdoNota("Qtde") & Chr(9) & Format(RdoNota("VLUNIT"), "###,###,###,#0.00") & Chr(9) & Format(RdoNota("VLUNIT2"), "###,###,###,#0.00")
                             RdoNota.MoveNext
                          Loop
                          
                          If grdItens.Rows > 2 And grdItens.TextMatrix(1, 0) = "" Then
                             grdItens.RemoveItem 1
                          End If
                          
                          RdoNota.Close
                     End If
                  End If
          Screen.MousePointer = 0
          
          'txtLjOri.SetFocus
          
          'txtCliente_LostFocus
          'EncontraTexto cmbCondPagto
       'End If
       
       txtSerie.Text = UCase(txtSerie.Text)
       
       Screen.MousePointer = 0
    Else
        MsgBox "Dados inválidos ou insuficientes.", vbExclamation, "Atenção"
        txtCFO.SetFocus
    End If

End Sub

Private Sub txtSubTotal_KeyPress(KeyAscii As Integer)

   VerteclaVirgula txtSubTotal, KeyAscii
   
End Sub

Private Sub txtSubTotal_LostFocus()

   txtSubTotal.Text = Format(txtSubTotal.Text, " ###,###,###,###,#0.00")
   
End Sub

Private Sub txtTotal_KeyPress(KeyAscii As Integer)

   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0
      GetAsyncKeyState vbKeyTab
      GetAsyncKeyState vbKeyTab
      fraDadosTotal.Visible = False
   End If
   
   VerteclaVirgula txtTotal, KeyAscii
End Sub

Private Sub txtTotal_LostFocus()
   If txtTotal.Text = "" Then
      txtTotal.Text = "0.00"
   Else
      txtTotal.Text = Format(txtTotal.Text, " ###,###,###,###,#0.00")
   End If
End Sub

Private Sub txtTotalItem_KeyPress(KeyAscii As Integer)

   VerteclaVirgula txtTotalItem, KeyAscii
   
End Sub

Private Sub cmbUFCli_LostFocus()

   If GetAsyncKeyState(vbKeyTab) <> 0 And cmbUFCli.Text <> "" Then
      If txtCFO.Text = 5102 Or txtCFO.Text = 6102 Then
         If Left(cmbCondPagto.Text, 1) <= 3 Then
            If txtCliente.Text = "888888" Then
               If GetAsyncKeyState(vbKeyTab) <> 0 And Trim(cmbUFCli.Text) = "" Then
                  MsgBox "Favor preencher o campo UF", vbInformation, "Atenção"
                  cmbUFCli.SetFocus
               Else
                  Exit Sub
               End If
            End If
         End If
      End If
      'txtUFCli.Text = UCase(txtUFCli.Text)
          
      If Len(UCase(cmbUFCli.Text)) <> 2 Or InStr("AC;AL;AM;AP;BA;DF;ES;GO;MG;MS;MT;PA;PB;PE;PI;PR;SC;SE;SP;RJ;RN;RO;RR;RS;TO", UCase(cmbUFCli.Text)) = 0 Then
         MsgBox "Estado inválido", vbInformation, "Informação"
         'txtUFCli.SelStart = 0
         'txtUFCli.SelLength = 2
         cmbUFCli.SetFocus
      End If
   
   End If
   
End Sub


Sub CalculaTotal()

   Dim ValQuant As Long
   Dim ValPreUnit As Double
   
   On Error Resume Next
   
   ValQuant = Val(Numeros(txtQuant.Text))
   ValPreUnit = CDbl(Format(txtPrecoUnit.Text, "0.0000"))
   
   If Err.Number <> 0 Then
      Err.Clear
      ValPreUnit = 0
   End If
   txtTotalItem.Text = Format(ValQuant * ValPreUnit, " ###,###,###,###,#0.00")
   
End Sub

Private Sub txtValorIPI_GotFocus()

   txtValorIPI.SelStart = 0
   txtValorIPI.SelLength = Len(txtValorIPI.Text)

End Sub

Private Sub txtValorIPI_KeyPress(KeyAscii As Integer)

   VerteclaVirgula txtValorIPI, KeyAscii
   
End Sub

Private Sub txtValorIPI_LostFocus()

   If txtValorIPI.Text = "" Then
      txtValorIPI.Text = "0.00"
   End If
   txtValorIPI.Text = Format(txtValorIPI.Text, " ###,###,###,###,#0.00")
    If txtTotal.Text <> "" Then
        If txtValorIPI.Text <> "" Then
           txtTotal.Text = Format(CDbl(txtTotal.Text) + CDbl(txtValorIPI.Text), " ###,###,###,###,#0.00")
           vlripi = CDbl(txtValorIPI.Text)
        Else
            txtTotal.Text = Format(CDbl(txtTotal.Text) - CDbl(vlripi), " ###,###,###,###,#0.00")
        End If
    End If
  
End Sub

Sub VerificaCodigo()

   txtCFO.Enabled = True
   cmbNaturezaOperacao.Enabled = True
   txtNF.Enabled = True
   txtSerie.Enabled = True
   txtLjOrigem.Enabled = True
   txtLjDestino.Enabled = True
   txtVendedor.Enabled = True
   cmbCondPagto.Enabled = True
   txtPagtoEnt.Enabled = True
   txtCliente.Enabled = True
   txtCGC.Enabled = True
   txtInscEst.Enabled = True
   txtRazao.Enabled = True
   txtEndereco.Enabled = True
   txtBairro.Enabled = True
   txtMunicipio.Enabled = True
   cmbUFCli.Enabled = True
   txtCEP.Enabled = True
   txtreferencia.Enabled = True
'   txtDescricao.Enabled = True
'   txtQuant.Enabled = True
'   txtPrecoUnit.Enabled = True
'   txtTotalItem.Enabled = True
'   txtSubTotal.Enabled = True
   txtDesconto.Enabled = True
   txtAliqIPI.Enabled = True
   txtValorIPI.Enabled = True
   txtFrete.Enabled = True
      
   Select Case CodigoOperAntigo
      Case 522
         lblCGCTransf.Caption = "CGC Transferência"
         lblLjDestino.Caption = "Lj. Destino"
         txtVendedor.Enabled = False
         txtOutroVendedor.Enabled = False
         cmbCondPagto.Enabled = False
         txtPagtoEnt.Enabled = False
         txtCliente.Enabled = False
         txtInscEst.Enabled = False
         txtRazao.Enabled = False
         txtEndereco.Enabled = False
         txtMunicipio.Enabled = False
         txtBairro.Enabled = False
         cmbUFCli.Enabled = False
         txtCEP.Enabled = False
         txtDesconto.Enabled = False
         txtAliqIPI.Enabled = True
         txtValorIPI.Enabled = True
         txtFrete.Enabled = False
         txtVendedor.Text = "999"
         txtOutroVendedor.Text = "0"
         txtAV.Enabled = False
         txtAV.Text = "0"
         txtLjDestino.Enabled = True
         txtLjDestino.SetFocus
         
      Case 512, 612, 519
         lblCGCTransf.Caption = "CGC"
         lblLjDestino.Caption = "Lj. Venda"
      
      Case 599, 699, 199, 299, 572, 672, 578, 678, 591, 691, 592, 595, 695
'         lblNatOper.Visible = False
'         cmbNaturezaOperacao.Visible = False
         txtVendedor.Text = "0"
         txtOutroVendedor.Text = "0"
         cmbCondPagto.Text = "0"
         txtPagtoEnt.Text = "0"
         txtLjDestino.Text = "0"
         txtDesconto.Text = "0"
         txtFrete.Text = "0"
         txtValorIPI.Text = "0"
         txtVendedor.Enabled = False
         txtOutroVendedor.Enabled = False
         cmbCondPagto.Enabled = False
         txtPagtoEnt.Enabled = False
         txtLjDestino.Enabled = False
         lblCGCTransf.Caption = "CGC"
         txtDesconto.Enabled = False
         txtFrete.Enabled = False
         txtCliente.Text = "888888"
         txtCliente.Enabled = False
         wTipoMovimentacao = 13
      
      Case 112, 212, 122, 143, 153, 163, 263, 172, 272, 173, 273, 174, 274, 191, 291, 192, 197, 297, 198
'         lblNatOper.Visible = False
'         cmbNaturezaOperacao.Visible = False
         txtVendedor.Text = "0"
         txtOutroVendedor.Text = "0"
         cmbCondPagto.Text = "0"
         txtPagtoEnt.Text = "0"
         txtLjDestino.Text = "0"
         txtDesconto.Text = "0"
         txtFrete.Text = "0"
         txtVendedor.Enabled = False
         txtOutroVendedor.Enabled = False
         cmbCondPagto.Enabled = False
         txtPagtoEnt.Enabled = False
         txtLjDestino.Enabled = False
         lblCGCTransf.Caption = "CGC"
         txtDesconto.Enabled = False
         txtFrete.Enabled = False
         txtCliente.Text = "888888"
         txtCliente.Enabled = False
         wTipoMovimentacao = 24
      
      Case 132, 232, 532, 632
         If CodigoOperNovo = 1202 Or CodigoOperNovo = 5202 Or CodigoOperNovo = 6202 Then
             txtVendedor.Enabled = False
             cmbCondPagto.Enabled = False
             txtPagtoEnt.Enabled = False
             txtDesconto.Enabled = False
             txtFrete.Enabled = False
             txtCliente.Text = "888888"
             txtCliente.Enabled = False
             txtLjDestino.Enabled = False
             txtOutroVendedor.Enabled = False
             txtAV.Enabled = False
             lblCGCTransf.Caption = "CGC"
             txtLjDestino.Text = txtLjOrigem.Text
             txtVendedor.Text = 0
             txtOutroVendedor.Text = 0
             txtAV.Text = 0
             txtFoneCli.Text = 0
             wPessoa = 1
         Else
             lblLjDestino.Caption = "Lj. Venda"
             lblCGCTransf.Caption = "CGC"
         End If
   End Select
   
End Sub

Private Function ObtemReferencias() As String

   Dim Retorno As String
   Dim Linha As Long
   Dim Maximo As Long
   
   Maximo = grdItens.Rows - 1
   
   Retorno = ""
   
   For Linha = 1 To Maximo Step 1
      Retorno = Retorno & ",'" & grdItens.TextMatrix(Linha, 0) & "'"
   Next Linha

   If Trim(Retorno) <> "" Then
      Retorno = Mid(Retorno, 2)
   End If

   ObtemReferencias = Retorno

End Function

Sub Gravar()
   
   Dim VerificaICMS As Double
   Dim rdoRegiao As rdoResultset
   Dim i As Long
   Dim y As Long
   Dim wQtdItem As Long
   
   SomaPeso = 0
   Peso = 0
   
   SQL = ""
   SQL = "Select UF_Regiao From Estados Where UF_Estado = '" & cmbUFCli.Text & "'"
   Set RSRegiao = Conexao.OpenResultset(SQL)
   
   WREGIAO = RSRegiao("UF_Regiao")
   
'   If txtBaseICMS.Text <> 0 Then
'      If txtAliqICMS.Text <> 0 Then
'         If txtValorICMS.Text <> 0 Then
'            VerificaICMS = CDbl(txtBaseICMS.Text) * CDbl(txtAliqICMS.Text) / 100
'            If Abs(VerificaICMS - CDbl(txtValorICMS.Text)) > 0.01 Then
'               Screen.MousePointer = 0
'               GravacaoOK = False
'               MsgBox "Campo valor ICMS incorreto.", vbExclamation, "Atenção"
'               Exit Sub
'            End If
'         End If
'      End If
'   End If
   
   WnumeroPed = ExtraiSeqPedido
   
   y = grdItens.Rows - 1
   
   wQtdItem = grdItens.Rows - 1
   
   For i = 1 To y Step 1
      Set rdoProduto = Conexao.OpenResultset("Select PR_PrecoVenda1 from Produto where PR_Referencia = '" & grdItens.TextMatrix(i, 0) & "'")

      Peso = (grdItens.RowData(i) / 10000) * grdItens.TextMatrix(i, 2)
      
      SomaPeso = SomaPeso + Peso
   Next i
   
   If mskDataPagto.Text = "__/__/____" Then
      mskDataPagto.Text = mskDataEmi.Text
   End If
   
   Select Case txtCFO.Text
      Case 5152
         wTipoMovimentacao = 12
         WTipoNota = "T"
         
         On Error Resume Next
         BeginTrans
'         SQL = "Insert into nfcapa (numeroped,dataemi,vendedor,VLRMERCADORIA,TOTALNOTA,DESCONTO, " _
'              & "SUBTOTAL,LOJAORIGEM,QTDITEM,TIPONOTA,CONDPAG,AV,CLIENTE,CODOPER,DATAPAG,PGENTRA, " _
'              & "LOJAT,PESOBR,PESOLQ,VALFRETE,FRETECOBR,OUTRALOJA,OUTROVEND,SERIE,UFCLIENTE, " _
'              & "NOMCLI,ENDCLI,CGCCLI,MUNICIPIOCLI,PESSOACLI,FONECLI,TM,INSCRICLI,BAIRROCLI, " _
'              & "CEPCLI,CARIMBO4,SituacaoEnvio,ValorTotalCodigoZero,TotalNotaAlternativa,ValorMercadoriaAlternativa,Carimbo3,CfoAux,LojaVenda,VendedorLojaVenda,PedCli,ECFNF,NF, RegiaoCli)" _
'              & "Values (" & WnumeroPed & ", '" & Format(mskDataEmi.Text, "mm/dd/yyyy") & "',0, " _
'              & "" & ConverteVirgula(Format(txtSubTotal.Text, "0.00")) & "," & ConverteVirgula(Format(txtTotal.Text, "0.00")) & "," & ConverteVirgula(Format(txtdesconto.text, "0.00")) & ", " _
'              & "" & ConverteVirgula(Format(txtSubTotal.Text, "0.00")) & ",'" & txtLjOrigem.Text & "'," & wQtdItem & ", " _
'              & "'T','0',0,999999, " & CodigoOperAntigo & ",'" & Format(mskDataEmi.Text, "dd/mm/yyyy") & "'," & ConverteVirgula(txtPagtoEnt.Text) & ", " _
'              & "'" & txtLjDestino.Text & "'," & SomaPeso & "," & SomaPeso & ", " _
'              & "" & ConverteVirgula(Format(txtFrete.Text, "0.00")) & "," & ConverteVirgula(Format(txtFrete.Text, "0.00")) & ",'" & txtLjDestino.Text & "',0, " _
'              & "'" & txtSerie.Text & "','" & cmbUFCli.Text & "','" & txtRazao.Text & "','" & txtEndereco.Text & "','" & txtCGC.Text & "','" & txtMunicipio.Text & "', " _
'              & "" & wPessoa & ",'" & txtFoneCli.Text & "','0','" & txtInscEst.Text & "','" & txtBairro.Text & "'," _
'              & "'" & txtCEP.Text & "',0,'A'," & ConverteVirgula(Format(0, "0.00")) & "," & ConverteVirgula(Format(0, "0.00")) & "," & ConverteVirgula(Format(0, "0.00")) & ",'0','" & txtCFO.Text & "','" & txtLjOrigem.Text & "','0', 0," & Val(glb_ECF) & "," & txtNF.Text & "," & WREGIAO & ")"
         
         SQL = ""
         SQL = "Insert into nfcapa (numeroped,dataemi,vendedor,VLRMERCADORIA,TOTALNOTA,DESCONTO, " _
              & "SUBTOTAL,LOJAORIGEM,QTDITEM,TIPONOTA,CONDPAG,AV,CLIENTE,CODOPER,DATAPAG,PGENTRA, " _
              & "LOJAT,PESOBR,PESOLQ,VALFRETE,FRETECOBR,OUTRALOJA,OUTROVEND,SERIE,UFCLIENTE, " _
              & "NOMCLI,ENDCLI,CGCCLI,MUNICIPIOCLI,PESSOACLI,FONECLI,TM,INSCRICLI,BAIRROCLI, " _
              & "CEPCLI,CARIMBO4,SituacaoEnvio,ValorTotalCodigoZero,TotalNotaAlternativa,ValorMercadoriaAlternativa,Carimbo3,CfoAux,LojaVenda,VendedorLojaVenda,PedCli,ECFNF,NF, RegiaoCli)" _
              & "Values (" & WnumeroPed & ", '" & Format(mskDataEmi.Text, "mm/dd/yyyy") & "',0, " _
              & "" & ConverteVirgula(Format(txtSubTotal.Text, "0.00")) & "," & ConverteVirgula(Format(txtTotal.Text, "0.00")) & "," & ConverteVirgula(Format(txtDesconto.Text, "0.00")) & ", " _
              & "" & ConverteVirgula(Format(txtSubTotal.Text, "0.00")) & ",'" & txtLjOrigem.Text & "'," & wQtdItem & ", " _
              & "'T','0',0,999999, " & CodigoOperAntigo & ",'" & Format(mskDataEmi.Text, "mm/dd/yyyy") & "'," & ConverteVirgula(txtPagtoEnt.Text) & ", " _
              & "'" & txtLjDestino.Text & "'," & SomaPeso & "," & SomaPeso & ", " _
              & "" & ConverteVirgula(Format(txtFrete.Text, "0.00")) & "," & ConverteVirgula(Format(txtFrete.Text, "0.00")) & ",'" & txtLjDestino.Text & "',0, " _
              & "'" & txtSerie.Text & "','" & cmbUFCli.Text & "','" & txtRazao.Text & "','" & txtEndereco.Text & "','" & txtCGC.Text & "','" & txtMunicipio.Text & "', " _
              & "1,'0','0','" & txtInscEst.Text & "','" & txtBairro.Text & "'," _
              & "'" & txtCEP.Text & "',0,'A'," & ConverteVirgula(Format(0, "0.00")) & "," & ConverteVirgula(Format(0, "0.00")) & "," & ConverteVirgula(Format(0, "0.00")) & ",'0','" & txtCFO.Text & "','" & txtLjOrigem.Text & "','0', 0," & Val(glb_ECF) & "," & txtNF.Text & "," & WREGIAO & ")"
             
         rdoCnLojaBach.Execute (SQL)
         If Err.Number = 0 Then
              CommitTrans
              wNfCapa = True
              GravacaoOK = True
              wPessoa = 1
         Else
              Rollback
              MsgBox "Não foi possível efetuar a gravação da capa", vbCritical, "Atenção!"
         End If
         
         WPERDESC = Val(txtDesconto / (txtSubTotal - txtDesconto)) * 100
         
         For i = 1 To grdItens.Rows - 1
            Set RSPegaProduto = rdoCnLojaBach.OpenResultset("Select PR_PrecoVenda1, PR_Linha, PR_Secao, PR_ICMPDV, PR_CodigoBarra from Produto where PR_Referencia = '" & grdItens.TextMatrix(i, 0) & "'")
            
            On Error Resume Next
            
            wVlDescRat = grdItens.TextMatrix(i, 3) * WPERDESC / 100
            
            BeginTrans
'            SQL = "Insert into nfitens(numeroped,dataemi,Referencia,Qtde,vlunit,vlunit2, " _
'                & "vltotitem,DESCRAT,ITEM,LINHA,SECAO,CSPROD,PLISTA,ICMS," _
'                & "ICMPDV,CODBARRA,NF,SERIE,CLIENTE,TIPONOTA,Vendedor,LojaOrigem,TipoMovimentacao,SituacaoEnvio,PrecoUnitAlternativa,ValorMercadoriaAlternativa,ReferenciaAlternativa,DescricaoAlternativa,SerieProd1,SerieProd2) " _
'                & "Values (" & WnumeroPed & ", '" & Format(mskDataEmi.Text, "mm/dd/yyyy") & "','" & grdItens.TextMatrix(i, 0) & "', " _
'                & "" & grdItens.TextMatrix(i, 2) & ", " & ConverteVirgula(Format(grdItens.TextMatrix(i, 3), "0.00")) & ", " & ConverteVirgula(Format(grdItens.TextMatrix(i, 4)- wvldescrat, "0.00")) & ", " _
'                & "" & ConverteVirgula(Format(grdItens.TextMatrix(i, 4), "0.00")) & "," & ConverteVirgula(Format(wvldescrat, "0.00")) & "," _
'                & "" & i & "," & RSPegaProduto("PR_Linha") & "," & RSPegaProduto("PR_Secao") & "," & WCSPROD & ", " _
'                & "" & ConverteVirgula(Format(RSPegaProduto("PR_PrecoVenda1"), "0.00")) & "," & ConverteVirgula(Format(txtAliqICMS.Text, "0.00")) & "," & ConverteVirgula(RSPegaProduto("PR_ICMPDV")) & ", " _
'                & "'" & RSPegaProduto("PR_CodigoBarra") & "'," & txtNF.Text & ", '" & txtSerie.Text & "'," & txtCliente.Text & ", " _
'                & "'T'," & numeros(txtVendedor.Text) & ",'" & txtLjOrigem.Text & "',12,'A'," & ConverteVirgula(0) & "," & ConverteVirgula(0) & ",'','','' , '')"
            
            SQL = ""
            SQL = "Insert into nfitens(numeroped,dataemi,Referencia,Qtde,vlunit,vlunit2, " _
                & "vltotitem,DESCRAT,ITEM,LINHA,SECAO,CSPROD,PLISTA,ICMS," _
                & "ICMPDV,CODBARRA,NF,SERIE,CLIENTE,TIPONOTA,Vendedor,LojaOrigem,TipoMovimentacao,SituacaoEnvio,PrecoUnitAlternativa,ValorMercadoriaAlternativa,ReferenciaAlternativa,DescricaoAlternativa,SerieProd1,SerieProd2) " _
                & "Values (" & WnumeroPed & ", '" & Format(mskDataEmi.Text, "mm/dd/yyyy") & "','" & grdItens.TextMatrix(i, 0) & "', " _
                & "" & grdItens.TextMatrix(i, 2) & ", " & ConverteVirgula(Format(grdItens.TextMatrix(i, 3), "0.00")) & ", " & ConverteVirgula(Format(grdItens.TextMatrix(i, 4) - wVlDescRat, "0.00")) & ", " _
                & "" & ConverteVirgula(Format(grdItens.TextMatrix(i, 4), "0.00")) & "," & ConverteVirgula(Format(wVlDescRat, "0.00")) & "," _
                & "" & i & "," & RSPegaProduto("PR_Linha") & "," & RSPegaProduto("PR_Secao") & "," & WCSPROD & ", " _
                & "" & ConverteVirgula(Format(RSPegaProduto("PR_PrecoVenda1"), "0.00")) & "," & ConverteVirgula(Format(0, "0.00")) & "," & ConverteVirgula(RSPegaProduto("PR_ICMPDV")) & ", " _
                & "'" & RSPegaProduto("PR_CodigoBarra") & "'," & txtNF.Text & ", '" & txtSerie.Text & "',0, " _
                & "'T',999,'" & txtLjOrigem.Text & "',12,'A'," & ConverteVirgula(0) & "," & ConverteVirgula(0) & ",'0','0','0' , '0')"
            
            rdoCnLojaBach.Execute (SQL)
            If Err.Number = 0 Then
               CommitTrans
            Else
               Rollback
               MsgBox "Não foi possível efetuar a gravação dos itens", vbCritical, "Atenção!"
            End If
            
         Next i
         
      Case 5102, 6102, 5202, 6202
         If txtCFO.Text = 5102 Or txtCFO.Text = 6102 Then
            wTipoMovimentacao = 11
            WTipoNota = "V"
         Else
            wTipoMovimentacao = 13
            cmbCondPagto.Text = 0
            WTipoNota = "S"
         End If
         
         BeginTrans
         If Left(cmbCondPagto.Text, 2) = 85 Then
'            SQL = "Insert into nfcapa (numeroped,dataemi,vendedor,VLRMERCADORIA,TOTALNOTA,DESCONTO, " _
'                 & "SUBTOTAL,LOJAORIGEM,QTDITEM,TIPONOTA,CONDPAG,AV,CLIENTE,CODOPER,DATAPAG,PGENTRA, " _
'                 & "LOJAT,PESOBR,PESOLQ,VALFRETE,FRETECOBR,OUTRALOJA,OUTROVEND,SERIE,UFCLIENTE, " _
'                 & "NOMCLI,ENDCLI,CGCCLI,MUNICIPIOCLI,PESSOACLI,FONECLI,TM,INSCRICLI,BAIRROCLI, " _
'                 & "CEPCLI,CARIMBO4,SituacaoEnvio,ValorTotalCodigoZero,TotalNotaAlternativa,ValorMercadoriaAlternativa,Carimbo3,CfoAux,LojaVenda,VendedorLojaVenda,PedCli,ECFNF,NF,RegiaoCLi)" _
'                 & "Values (" & WnumeroPed & ", '" & Format(mskDataEmi.Text, "mm/dd/yyyy") & "'," & Numeros(txtVendedor.Text) & ", " _
'                 & "" & ConverteVirgula(Format(txtSubTotal.Text, "0.00")) & "," & ConverteVirgula(Format(txtTotal.Text, "0.00")) & "," & ConverteVirgula(Format(txtdesconto.text, "0.00")) & ", " _
'                 & "" & ConverteVirgula(Format(txtSubTotal.Text, "0.00")) & ",'" & txtLjOrigem.Text & "'," & wQtdItem & ", " _
'                 & "'V','" & Numeros(cmbCondPagto.Text) & "'," & txtAV.Text & "," & txtCliente.Text & ", " & CodigoOperAntigo & ",'" & Format(mskDataPagto.Text, "dd/mm/yyyy") & "'," & ConverteVirgula(txtPagtoEnt.Text) & ", " _
'                 & "'" & txtLjDestino.Text & "'," & ConverteVirgula(SomaPeso) & "," & ConverteVirgula(SomaPeso) & ", " _
'                 & "" & ConverteVirgula(Format(txtFrete.Text, "0.00")) & "," & ConverteVirgula(Format(txtFrete.Text, "0.00")) & ",'" & txtLjDestino.Text & "'," & Numeros(txtOutroVendedor.Text) & ", " _
'                 & "'" & txtSerie.Text & "','" & cmbUFCli.Text & "','" & txtRazao.Text & "','" & txtEndereco.Text & "','" & txtCGC.Text & "','" & txtMunicipio.Text & "', " _
'                 & "" & wPessoa & ",'" & txtFoneCli.Text & "','0','" & txtInscEst.Text & "','" & txtBairro.Text & "'," _
'                 & "'" & txtCEP.Text & "',0,'A'," & ConverteVirgula(Format(0, "0.00")) & "," & ConverteVirgula(Format(0, "0.00")) & "," & ConverteVirgula(Format(0, "0.00")) & ",'0','" & txtCFO.Text & "','" & txtLjOrigem.Text & "','0', 0," & Val(glb_ECF) & "," & txtNF.Text & "," & WREGIAO & ")"
            
            SQL = ""
            SQL = "Insert into nfcapa (numeroped,dataemi,vendedor,VLRMERCADORIA,TOTALNOTA,DESCONTO, " _
                 & "SUBTOTAL,LOJAORIGEM,QTDITEM,TIPONOTA,CONDPAG,AV,CLIENTE,CODOPER,DATAPAG,PGENTRA, " _
                 & "LOJAT,PESOBR,PESOLQ,VALFRETE,FRETECOBR,OUTRALOJA,OUTROVEND,SERIE,UFCLIENTE, " _
                 & "NOMCLI,ENDCLI,CGCCLI,MUNICIPIOCLI,PESSOACLI,FONECLI,TM,INSCRICLI,BAIRROCLI, " _
                 & "CEPCLI,CARIMBO4,SituacaoEnvio,ValorTotalCodigoZero,TotalNotaAlternativa,ValorMercadoriaAlternativa,Carimbo3,CfoAux,LojaVenda,VendedorLojaVenda,PedCli,ECFNF,NF,RegiaoCLi)" _
                 & "Values (" & WnumeroPed & ", '" & Format(mskDataEmi.Text, "mm/dd/yyyy") & "'," & Numeros(txtVendedor.Text) & ", " _
                 & "" & ConverteVirgula(Format(txtSubTotal.Text, "0.00")) & "," & ConverteVirgula(Format(txtTotal.Text, "0.00")) & "," & ConverteVirgula(Format(txtDesconto.Text, "0.00")) & ", " _
                 & "" & ConverteVirgula(Format(txtSubTotal.Text, "0.00")) & ",'" & txtLjOrigem.Text & "'," & wQtdItem & ", " _
                 & "'V','" & Numeros(cmbCondPagto.Text) & "'," & txtAV.Text & "," & txtCliente.Text & ", " & CodigoOperAntigo & ",'" & Format(mskDataPagto.Text, "dd/mm/yyyy") & "'," & ConverteVirgula(txtPagtoEnt.Text) & ", " _
                 & "'" & txtLjDestino.Text & "'," & ConverteVirgula(SomaPeso) & "," & ConverteVirgula(SomaPeso) & ", " _
                 & "" & ConverteVirgula(Format(txtFrete.Text, "0.00")) & "," & ConverteVirgula(Format(txtFrete.Text, "0.00")) & ",'" & txtLjDestino.Text & "'," & Numeros(txtOutroVendedor.Text) & ", " _
                 & "'" & txtSerie.Text & "','" & cmbUFCli.Text & "','" & txtRazao.Text & "','" & txtEndereco.Text & "','" & txtCGC.Text & "','" & txtMunicipio.Text & "', " _
                 & "" & wPessoa & ",'" & txtFoneCli.Text & "','0','" & txtInscEst.Text & "','" & txtBairro.Text & "'," _
                 & "'" & txtCEP.Text & "',0,'A'," & ConverteVirgula(Format(0, "0.00")) & "," & ConverteVirgula(Format(0, "0.00")) & "," & ConverteVirgula(Format(0, "0.00")) & ",'0','" & txtCFO.Text & "','" & txtLjOrigem.Text & "','0', 0," & Val(glb_ECF) & "," & txtNF.Text & "," & WREGIAO & ")"
                
         Else
            SQL = ""
            SQL = "Insert into nfcapa (numeroped,dataemi,vendedor,VLRMERCADORIA,TOTALNOTA,DESCONTO, " _
                 & "SUBTOTAL,LOJAORIGEM,QTDITEM,TIPONOTA,CONDPAG,AV,CLIENTE,CODOPER,DATAPAG,PGENTRA, " _
                 & "LOJAT,PESOBR,PESOLQ,VALFRETE,FRETECOBR,OUTRALOJA,OUTROVEND,SERIE,UFCLIENTE, " _
                 & "NOMCLI,ENDCLI,CGCCLI,MUNICIPIOCLI,PESSOACLI,FONECLI,TM,INSCRICLI,BAIRROCLI, " _
                 & "CEPCLI,CARIMBO4,SituacaoEnvio,ValorTotalCodigoZero,TotalNotaAlternativa,ValorMercadoriaAlternativa,Carimbo3,CfoAux,LojaVenda,VendedorLojaVenda,PedCli,ECFNF,NF,RegiaoCli)" _
                 & "Values (" & WnumeroPed & ", '" & Format(mskDataEmi.Text, "mm/dd/yyyy") & "'," & Numeros(txtVendedor.Text) & ", " _
                 & "" & ConverteVirgula(Format(txtSubTotal.Text, "0.00")) & "," & ConverteVirgula(Format(txtTotal.Text, "0.00")) & "," & ConverteVirgula(Format(txtDesconto.Text, "0.00")) & ", " _
                 & "" & ConverteVirgula(Format(txtSubTotal.Text, "0.00")) & ",'" & txtLjOrigem.Text & "'," & wQtdItem & ", " _
                 & "'" & WTipoNota & "'," & Mid(cmbCondPagto.Text, 1, 2) & "," & txtAV.Text & "," & txtCliente.Text & ", " & CodigoOperAntigo & ",'" & Format(mskDataPagto.Text, "mm/dd/yyyy") & "'," & ConverteVirgula(txtPagtoEnt.Text) & ", " _
                 & "'" & txtLjDestino.Text & "'," & ConverteVirgula(SomaPeso) & "," & ConverteVirgula(SomaPeso) & ", " _
                 & "" & ConverteVirgula(Format(txtFrete.Text, "0.00")) & "," & ConverteVirgula(Format(txtFrete.Text, "0.00")) & ",'" & txtLjDestino.Text & "'," & Numeros(txtOutroVendedor.Text) & ", " _
                 & "'" & txtSerie.Text & "','" & cmbUFCli.Text & "','" & txtRazao.Text & "','" & txtEndereco.Text & "','" & txtCGC.Text & "','" & txtMunicipio.Text & "', " _
                 & "" & wPessoa & ",'" & txtFoneCli.Text & "','0','" & txtInscEst.Text & "','" & txtBairro.Text & "'," _
                 & "'" & txtCEP.Text & "',0,'A'," & ConverteVirgula(Format(0, "0.00")) & "," & ConverteVirgula(Format(0, "0.00")) & "," & ConverteVirgula(Format(0, "0.00")) & ",'0','" & txtCFO.Text & "','" & txtLjDestino.Text & "'," & Numeros(txtOutroVendedor.Text) & ", 0," & Val(glb_ECF) & "," & txtNF.Text & "," & WREGIAO & ")"
         End If
         
         rdoCnLojaBach.Execute (SQL)
         If Err.Number = 0 Then
              CommitTrans
              wNfCapa = True
              GravacaoOK = True
         Else
              Rollback
              MsgBox "Não foi possível efetuar a gravação da capa", vbCritical, "Atenção!"
         End If
         
         WPERDESC = CDbl(txtDesconto) / CDbl(txtSubTotal) * 100
         
         For i = 1 To grdItens.Rows - 1
            Set RSPegaProduto = rdoCnLojaBach.OpenResultset("Select PR_PrecoVenda1, PR_Linha, PR_Secao, PR_ICMPDV, PR_CodigoBarra from Produto where PR_Referencia = '" & grdItens.TextMatrix(i, 0) & "'")
            
            wVlDescRat = grdItens.TextMatrix(i, 4) * WPERDESC / 100

'            On Error Resume Next
            BeginTrans
'            SQL = "Insert into nfitens(numeroped,dataemi,Referencia,Qtde,vlunit,vlunit2, " _
'                & "vltotitem,DESCRAT,ITEM,LINHA,SECAO,CSPROD,PLISTA,ICMS," _
'                & "ICMPDV,CODBARRA,NF,SERIE,CLIENTE,TIPONOTA,Vendedor,LojaOrigem,TipoMovimentacao,SituacaoEnvio,PrecoUnitAlternativa,ValorMercadoriaAlternativa,ReferenciaAlternativa,DescricaoAlternativa,SerieProd1,SerieProd2) " _
'                & "Values (" & WnumeroPed & ", '" & Format(mskDataEmi.Text, "mm/dd/yyyy") & "','" & grdItens.TextMatrix(i, 0) & "', " _
'                & "" & grdItens.TextMatrix(i, 2) & ", " & ConverteVirgula(Format(grdItens.TextMatrix(i, 3), "0.00")) & ", " & ConverteVirgula(Format(grdItens.TextMatrix(i, 4)- wvldescrat, "0.00")) & ", " _
'                & "" & ConverteVirgula(Format(grdItens.TextMatrix(i, 4), "0.00")) & "," & ConverteVirgula(Format(wvldescrat, "0.00")) & "," _
'                & "" & i & "," & RSPegaProduto("PR_Linha") & "," & RSPegaProduto("PR_Secao") & "," & WCSPROD & ", " _
'                & "" & ConverteVirgula(Format(RSPegaProduto("PR_PrecoVenda1"), "0.00")) & "," & ConverteVirgula(Format(txtAliqICMS.Text, "0.00")) & "," & ConverteVirgula(RSPegaProduto("PR_ICMPDV")) & ", " _
'                & "'" & RSPegaProduto("PR_CodigoBarra") & "'," & txtNF.Text & ", '" & txtSerie.Text & "'," & txtCliente.Text & ", " _
'                & "'V'," & numeros(txtVendedor.Text) & ",'" & txtLjOrigem.Text & "',11,'A'," & ConverteVirgula(0) & "," & ConverteVirgula(0) & ",'','','' , '')"
            
            SQL = ""
            SQL = "Insert into nfitens(numeroped,dataemi,Referencia,Qtde,vlunit,vlunit2, " _
                & "vltotitem,DESCRAT,ITEM,LINHA,SECAO,CSPROD,PLISTA,ICMS," _
                & "ICMPDV,CODBARRA,NF,SERIE,CLIENTE,TIPONOTA,Vendedor,LojaOrigem,TipoMovimentacao,SituacaoEnvio,PrecoUnitAlternativa,ValorMercadoriaAlternativa,ReferenciaAlternativa,DescricaoAlternativa,SerieProd1,SerieProd2) " _
                & "Values (" & WnumeroPed & ", '" & Format(mskDataEmi.Text, "mm/dd/yyyy") & "','" & grdItens.TextMatrix(i, 0) & "', " _
                & "" & grdItens.TextMatrix(i, 2) & ", " & ConverteVirgula(Format(grdItens.TextMatrix(i, 3), "0.00")) & ", " & ConverteVirgula(Format(grdItens.TextMatrix(i, 4) - wVlDescRat, "0.00")) & ", " _
                & "" & ConverteVirgula(Format(grdItens.TextMatrix(i, 4), "0.00")) & "," & ConverteVirgula(Format(wVlDescRat, "0.00")) & "," _
                & "" & i & "," & RSPegaProduto("PR_Linha") & "," & RSPegaProduto("PR_Secao") & "," & WCSPROD & ", " _
                & "" & ConverteVirgula(Format(RSPegaProduto("PR_PrecoVenda1"), "0.00")) & "," & ConverteVirgula(Format(0, "0.00")) & "," & ConverteVirgula(RSPegaProduto("PR_ICMPDV")) & ", " _
                & "'" & RSPegaProduto("PR_CodigoBarra") & "'," & txtNF.Text & ", '" & txtSerie.Text & "'," & txtCliente.Text & ", " _
                & "'" & WTipoNota & "'," & Numeros(txtVendedor.Text) & ",'" & txtLjOrigem.Text & "'," & wTipoMovimentacao & ",'A'," & ConverteVirgula(0) & "," & ConverteVirgula(0) & ",'0','0','0' , '0')"
            
            rdoCnLojaBach.Execute (SQL)
            If Err.Number = 0 Then
               CommitTrans
            Else
               Rollback
               MsgBox "Não foi possível efetuar a gravação dos itens", vbCritical, "Atenção!"
            End If
            
         Next i
         
      Case 5910, 6910, 5911, 6911, 5912, 6912, 5913, 6913, 5914, 6914, 5915, 6915, 5917, 6917, 5918, 6918, 5949, 6949
            wTipoMovimentacao = 13
            cmbCondPagto.Text = 0
            WTipoNota = "S"
            wPessoa = 1
            BeginTrans
            SQL = ""
            SQL = "Insert into nfcapa (numeroped,dataemi,vendedor,VLRMERCADORIA,TOTALNOTA,DESCONTO, " _
                 & "SUBTOTAL,LOJAORIGEM,QTDITEM,TIPONOTA,CONDPAG,AV,CLIENTE,CODOPER,DATAPAG,PGENTRA, " _
                 & "LOJAT,PESOBR,PESOLQ,VALFRETE,FRETECOBR,OUTRALOJA,OUTROVEND,SERIE,TOTALIPI,UFCLIENTE, " _
                 & "NOMCLI,ENDCLI,CGCCLI,MUNICIPIOCLI,PESSOACLI,FONECLI,TM,INSCRICLI,BAIRROCLI, " _
                 & "CEPCLI,CARIMBO4,SituacaoEnvio,ValorTotalCodigoZero,TotalNotaAlternativa,ValorMercadoriaAlternativa,Carimbo3,CfoAux,LojaVenda,VendedorLojaVenda,PedCli,ECFNF,NF,RegiaoCLi)" _
                 & "Values (" & WnumeroPed & ", '" & Format(mskDataEmi.Text, "mm/dd/yyyy") & "'," & Numeros(txtVendedor.Text) & ", " _
                 & "" & ConverteVirgula(Format(txtSubTotal.Text, "0.00")) & "," & ConverteVirgula(Format(txtTotal.Text, "0.00")) & "," & ConverteVirgula(Format(txtDesconto.Text, "0.00")) & ", " _
                 & "" & ConverteVirgula(Format(txtSubTotal.Text, "0.00")) & ",'" & txtLjOrigem.Text & "'," & wQtdItem & ", " _
                 & "'S','" & Numeros(cmbCondPagto.Text) & "'," & txtAV.Text & "," & txtCliente.Text & ", " & CodigoOperAntigo & ",'" & Format(mskDataPagto.Text, "mm/dd/yyyy") & "'," & ConverteVirgula(txtPagtoEnt.Text) & ", " _
                 & "'" & txtLjDestino.Text & "'," & SomaPeso & "," & SomaPeso & ", " _
                 & "" & ConverteVirgula(Format(txtFrete.Text, "0.00")) & "," & ConverteVirgula(Format(txtFrete.Text, "0.00")) & ",'" & txtLjDestino.Text & "'," & Numeros(txtOutroVendedor.Text) & ", " _
                 & "'" & txtSerie.Text & "'," & ConverteVirgula(Format(txtValorIPI.Text, "0.00")) & ",'" & cmbUFCli.Text & "','" & txtRazao.Text & "','" & txtEndereco.Text & "','" & txtCGC.Text & "','" & txtMunicipio.Text & "', " _
                 & "" & wPessoa & ",'0','0','" & txtInscEst.Text & "','" & txtBairro.Text & "'," _
                 & "'" & txtCEP.Text & "',0,'A'," & ConverteVirgula(Format(0, "0.00")) & "," & ConverteVirgula(Format(0, "0.00")) & "," & ConverteVirgula(Format(0, "0.00")) & ",'0','" & txtCFO.Text & "','" & txtLjOrigem.Text & "','0', 0," & Val(glb_ECF) & "," & txtNF.Text & "," & WREGIAO & ")"
            
            rdoCnLojaBach.Execute (SQL)
            If Err.Number = 0 Then
                 CommitTrans
                 wNfCapa = True
                 GravacaoOK = True
            Else
                 Rollback
                 MsgBox "Não foi possível efetuar a gravação da capa", vbCritical, "Atenção!"
            End If
            
            WPERDESC = Val(txtDesconto / (txtSubTotal - txtDesconto)) * 100
            
            For i = 1 To grdItens.Rows - 1
                Set RSPegaProduto = rdoCnLojaBach.OpenResultset("Select PR_PrecoVenda1, PR_Linha, PR_Secao, PR_ICMPDV, PR_CodigoBarra from Produto where PR_Referencia = '" & grdItens.TextMatrix(i, 0) & "'")
                
                wVlDescRat = grdItens.TextMatrix(i, 3) * WPERDESC / 100
                
                On Error Resume Next
                BeginTrans
    '            SQL = "Insert into nfitens(numeroped,dataemi,Referencia,Qtde,vlunit,vlunit2, " _
    '                & "vltotitem,DESCRAT,ITEM,LINHA,SECAO,CSPROD,PLISTA,ICMS," _
    '                & "ICMPDV,CODBARRA,NF,SERIE,CLIENTE,TIPONOTA,Vendedor,LojaOrigem,TipoMovimentacao,SituacaoEnvio,PrecoUnitAlternativa,ValorMercadoriaAlternativa,ReferenciaAlternativa,DescricaoAlternativa,SerieProd1,SerieProd2) " _
    '                & "Values (" & WnumeroPed & ", '" & Format(mskDataEmi.Text, "mm/dd/yyyy") & "','" & grdItens.TextMatrix(i, 0) & "', " _
    '                & "" & grdItens.TextMatrix(i, 2) & ", " & ConverteVirgula(Format(grdItens.TextMatrix(i, 3), "0.00")) & ", " & ConverteVirgula(Format(grdItens.TextMatrix(i, 4)- wVlDescRat, "0.00")) & ", " _
    '                & "" & ConverteVirgula(Format(grdItens.TextMatrix(i, 4), "0.00")) & "," & ConverteVirgula(Format(wvldescrat, "0.00")) & "," _
    '                & "" & i & "," & RSPegaProduto("PR_Linha") & "," & RSPegaProduto("PR_Secao") & "," & WCSPROD & ", " _
    '                & "" & ConverteVirgula(Format(RSPegaProduto("PR_PrecoVenda1"), "0.00")) & "," & ConverteVirgula(Format(txtAliqICMS.Text, "0.00")) & "," & ConverteVirgula(RSPegaProduto("PR_ICMPDV")) & ", " _
    '                & "'" & RSPegaProduto("PR_CodigoBarra") & "'," & txtNF.Text & ", '" & txtSerie.Text & "'," & txtCliente.Text & ", " _
    '                & "'S'," & numeros(txtVendedor.Text) & ",'" & txtLjOrigem.Text & "',13,'A'," & ConverteVirgula(0) & "," & ConverteVirgula(0) & ",'','','' , '')"
                
                SQL = ""
                SQL = "Insert into nfitens(numeroped,dataemi,Referencia,Qtde,vlunit,vlunit2, " _
                    & "vltotitem,DESCRAT,ITEM,LINHA,SECAO,CSPROD,PLISTA,ICMS," _
                    & "ICMPDV,CODBARRA,NF,SERIE,CLIENTE,TIPONOTA,Vendedor,LojaOrigem,TipoMovimentacao,SituacaoEnvio,PrecoUnitAlternativa,ValorMercadoriaAlternativa,ReferenciaAlternativa,DescricaoAlternativa,SerieProd1,SerieProd2) " _
                    & "Values (" & WnumeroPed & ", '" & Format(mskDataEmi.Text, "mm/dd/yyyy") & "','" & grdItens.TextMatrix(i, 0) & "', " _
                    & "" & grdItens.TextMatrix(i, 2) & ", " & ConverteVirgula(Format(grdItens.TextMatrix(i, 3), "0.00")) & ", " & ConverteVirgula(Format(grdItens.TextMatrix(i, 4) - wVlDescRat, "0.00")) & ", " _
                    & "" & ConverteVirgula(Format(grdItens.TextMatrix(i, 4), "0.00")) & "," & ConverteVirgula(Format(wVlDescRat, "0.00")) & "," _
                    & "" & i & "," & RSPegaProduto("PR_Linha") & "," & RSPegaProduto("PR_Secao") & "," & WCSPROD & ", " _
                    & "" & ConverteVirgula(Format(RSPegaProduto("PR_PrecoVenda1"), "0.00")) & "," & ConverteVirgula(Format(0, "0.00")) & "," & ConverteVirgula(RSPegaProduto("PR_ICMPDV")) & ", " _
                    & "'" & RSPegaProduto("PR_CodigoBarra") & "'," & txtNF.Text & ", '" & txtSerie.Text & "'," & txtCliente.Text & ", " _
                    & "'S'," & Numeros(txtVendedor.Text) & ",'" & txtLjOrigem.Text & "',13,'A'," & ConverteVirgula(0) & "," & ConverteVirgula(0) & ",'0','0','0' , '0')"
                
                rdoCnLojaBach.Execute (SQL)
                If Err.Number = 0 Then
                   CommitTrans
                Else
                   Rollback
                   MsgBox "Não foi possível efetuar a gravação dos itens", vbCritical, "Atenção!"
                End If
            Next i
      Case 1910, 2910, 1911, 2911, 1912, 2912, 1913, 2913, 1914, 2914, 1916, 2916, 1917, 2917, 1918, 2918, 1949, 2949, 1202, 2202
            wTipoMovimentacao = 24
            cmbCondPagto.Text = 0
            WTipoNota = "E"
            wPessoa = 1
            BeginTrans
           
            SQL = ""
            SQL = "Insert into nfcapa (numeroped,dataemi,vendedor,VLRMERCADORIA,TOTALNOTA,DESCONTO, " _
                 & "SUBTOTAL,LOJAORIGEM,QTDITEM,TIPONOTA,CONDPAG,AV,CLIENTE,CODOPER,DATAPAG,PGENTRA, " _
                 & "LOJAT,PESOBR,PESOLQ,VALFRETE,FRETECOBR,OUTRALOJA,OUTROVEND,SERIE,UFCLIENTE, " _
                 & "NOMCLI,ENDCLI,CGCCLI,MUNICIPIOCLI,PESSOACLI,FONECLI,TM,INSCRICLI,BAIRROCLI, " _
                 & "CEPCLI,CARIMBO4,SituacaoEnvio,ValorTotalCodigoZero,TotalNotaAlternativa,ValorMercadoriaAlternativa,Carimbo3,CfoAux,LojaVenda,VendedorLojaVenda,PedCli,ECFNF,NF,RegiaoCli)" _
                 & "Values (" & WnumeroPed & ", '" & Format(mskDataEmi.Text, "mm/dd/yyyy") & "'," & Numeros(txtVendedor.Text) & ", " _
                 & "" & ConverteVirgula(Format(txtSubTotal.Text, "0.00")) & "," & ConverteVirgula(Format(txtTotal.Text, "0.00")) & "," & ConverteVirgula(Format(txtDesconto.Text, "0.00")) & ", " _
                 & "" & ConverteVirgula(Format(txtSubTotal.Text, "0.00")) & ",'" & txtLjOrigem.Text & "'," & wQtdItem & ", " _
                 & "'" & WTipoNota & "','" & Numeros(cmbCondPagto.Text) & "'," & txtAV.Text & "," & txtCliente.Text & ", " & CodigoOperAntigo & ",'" & Format(mskDataPagto.Text, "mm/dd/yyyy") & "'," & ConverteVirgula(txtPagtoEnt.Text) & ", " _
                 & "'" & txtLjDestino.Text & "'," & ConverteVirgula(SomaPeso) & "," & ConverteVirgula(SomaPeso) & ", " _
                 & "" & ConverteVirgula(Format(txtFrete.Text, "0.00")) & "," & ConverteVirgula(Format(txtFrete.Text, "0.00")) & ",'" & txtLjDestino.Text & "'," & Numeros(txtOutroVendedor.Text) & ", " _
                 & "'" & txtSerie.Text & "','" & cmbUFCli.Text & "','" & txtRazao.Text & "','" & txtEndereco.Text & "','" & txtCGC.Text & "','" & txtMunicipio.Text & "', " _
                 & "" & wPessoa & ",'" & txtFoneCli.Text & "','0','" & txtInscEst.Text & "','" & txtBairro.Text & "'," _
                 & "'" & txtCEP.Text & "',0,'A'," & ConverteVirgula(Format(0, "0.00")) & "," & ConverteVirgula(Format(0, "0.00")) & "," & ConverteVirgula(Format(0, "0.00")) & ",'0','" & txtCFO.Text & "','" & txtLjDestino.Text & "'," & Numeros(txtOutroVendedor.Text) & ", 0," & Val(glb_ECF) & "," & txtNF.Text & "," & WREGIAO & ")"
            
'            SQL = "Insert into nfcapa (numeroped,dataemi,vendedor,VLRMERCADORIA,TOTALNOTA,DESCONTO, " _
'                 & "SUBTOTAL,LOJAORIGEM,QTDITEM,TIPONOTA,CONDPAG,AV,CLIENTE,CODOPER,DATAPAG,PGENTRA, " _
'                 & "LOJAT,PESOBR,PESOLQ,VALFRETE,FRETECOBR,OUTRALOJA,OUTROVEND,SERIE,TOTALIPI,UFCLIENTE, " _
'                 & "NOMCLI,ENDCLI,CGCCLI,MUNICIPIOCLI,PESSOACLI,FONECLI,TM,INSCRICLI,BAIRROCLI, " _
'                 & "CEPCLI,CARIMBO4,SituacaoEnvio,ValorTotalCodigoZero,TotalNotaAlternativa,ValorMercadoriaAlternativa,Carimbo3,CfoAux,LojaVenda,VendedorLojaVenda,PedCli,ECFNF,NF,RegiaoCli)" _
'                 & "Values (" & WnumeroPed & ", '" & Format(mskDataEmi.Text, "mm/dd/yyyy") & "'," & Numeros(txtVendedor.Text) & ", " _
'                 & "" & ConverteVirgula(Format(txtSubTotal.Text, "0.00")) & "," & ConverteVirgula(Format(txtTotal.Text, "0.00")) & "," & ConverteVirgula(Format(txtDesconto.Text, "0.00")) & ", " _
'                 & "" & ConverteVirgula(Format(txtSubTotal.Text, "0.00")) & ",'" & txtLjOrigem.Text & "'," & wQtdItem & ", " _
'                 & "'E'," & Numeros(cmbCondPagto.Text) & "," & txtAV.Text & "," & txtCliente.Text & ", " & CodigoOperAntigo & ",'" & Format(mskDataPagto.Text, "dd/mm/yyyy") & "'," & ConverteVirgula(txtPagtoEnt.Text) & ", " _
'                 & "'" & txtLjDestino.Text & "'," & SomaPeso & "," & SomaPeso & ", " _
'                 & "" & ConverteVirgula(Format(txtFrete.Text, "0.00")) & "," & ConverteVirgula(Format(txtFrete.Text, "0.00")) & ",'" & txtLjDestino.Text & "'," & Numeros(txtOutroVendedor.Text) & ", " _
'                 & "'" & txtSerie.Text & "'," & ConverteVirgula(Format(txtValorIPI.Text, "0.00")) & ",'" & cmbUFCli.Text & "','" & txtRazao.Text & "','" & txtEndereco.Text & "','" & txtCGC.Text & "','" & txtMunicipio.Text & "', " _
'                 & "" & wPessoa & ",'0','0','" & txtInscEst.Text & "','" & txtBairro.Text & "'," _
'                 & "'" & txtCEP.Text & "',0,'A'," & ConverteVirgula(Format(0, "0.00")) & "," & ConverteVirgula(Format(0, "0.00")) & "," & ConverteVirgula(Format(0, "0.00")) & ",'0','" & txtCFO.Text & "','" & txtLjOrigem.Text & "','0', 0," & Val(glb_ECF) & "," & txtNF.Text & "," & WREGIAO & ")"
            
            rdoCnLojaBach.Execute (SQL)
            If Err.Number = 0 Then
                 CommitTrans
                 wNfCapa = True
                 GravacaoOK = True
            Else
                 Rollback
                 MsgBox "Não foi possível efetuar a gravação da capa", vbCritical, "Atenção!"
            End If
            
            WPERDESC = Val(txtDesconto / (txtSubTotal - txtDesconto)) * 100
            
            For i = 1 To grdItens.Rows - 1
                Set RSPegaProduto = rdoCnLojaBach.OpenResultset("Select PR_PrecoVenda1, PR_Linha, PR_Secao, PR_ICMPDV, PR_CodigoBarra from Produto where PR_Referencia = '" & grdItens.TextMatrix(i, 0) & "'")
                
                wVlDescRat = grdItens.TextMatrix(i, 3) * WPERDESC / 100
                
                On Error Resume Next
                BeginTrans
    '            SQL = "Insert into nfitens(numeroped,dataemi,Referencia,Qtde,vlunit,vlunit2, " _
    '                & "vltotitem,DESCRAT,ITEM,LINHA,SECAO,CSPROD,PLISTA,ICMS," _
    '                & "ICMPDV,CODBARRA,NF,SERIE,CLIENTE,TIPONOTA,Vendedor,LojaOrigem,TipoMovimentacao,SituacaoEnvio,PrecoUnitAlternativa,ValorMercadoriaAlternativa,ReferenciaAlternativa,DescricaoAlternativa,SerieProd1,SerieProd2) " _
    '                & "Values (" & WnumeroPed & ", '" & Format(mskDataEmi.Text, "mm/dd/yyyy") & "','" & grdItens.TextMatrix(i, 0) & "', " _
    '                & "" & grdItens.TextMatrix(i, 2) & ", " & ConverteVirgula(Format(grdItens.TextMatrix(i, 3), "0.00")) & ", " & ConverteVirgula(Format(grdItens.TextMatrix(i, 2), "0.00")) & ", " _
    '                & "" & ConverteVirgula(Format(grdItens.TextMatrix(i, 4), "0.00")) & "," & ConverteVirgula(Format(wvldescrat, "0.00")) & "," _
    '                & "" & i & "," & RSPegaProduto("PR_Linha") & "," & RSPegaProduto("PR_Secao") & "," & WCSPROD & ", " _
    '                & "" & ConverteVirgula(Format(RSPegaProduto("PR_PrecoVenda1"), "0.00")) & "," & ConverteVirgula(Format(txtAliqICMS.Text, "0.00")) & "," & ConverteVirgula(RSPegaProduto("PR_ICMPDV")) & ", " _
    '                & "'" & RSPegaProduto("PR_CodigoBarra") & "'," & txtNF.Text & ", '" & txtSerie.Text & "'," & txtCliente.Text & ", " _
    '                & "'E'," & numeros(txtVendedor.Text) & ",'" & txtLjOrigem.Text & "',24,'A'," & ConverteVirgula(0) & "," & ConverteVirgula(0) & ",'','','' , '')"
                
                SQL = ""
                SQL = "Insert into nfitens(numeroped,dataemi,Referencia,Qtde,vlunit,vlunit2, " _
                    & "vltotitem,DESCRAT,ITEM,LINHA,SECAO,CSPROD,PLISTA,ICMS," _
                    & "ICMPDV,CODBARRA,NF,SERIE,CLIENTE,TIPONOTA,Vendedor,LojaOrigem,TipoMovimentacao,SituacaoEnvio,PrecoUnitAlternativa,ValorMercadoriaAlternativa,ReferenciaAlternativa,DescricaoAlternativa,SerieProd1,SerieProd2) " _
                    & "Values (" & WnumeroPed & ", '" & Format(mskDataEmi.Text, "mm/dd/yyyy") & "','" & grdItens.TextMatrix(i, 0) & "', " _
                    & "" & grdItens.TextMatrix(i, 2) & ", " & ConverteVirgula(Format(grdItens.TextMatrix(i, 3), "0.00")) & ", " & ConverteVirgula(Format(grdItens.TextMatrix(i, 4) - wVlDescRat, "0.00")) & ", " _
                    & "" & ConverteVirgula(Format(grdItens.TextMatrix(i, 4), "0.00")) & "," & ConverteVirgula(Format(wVlDescRat, "0.00")) & "," _
                    & "" & i & "," & RSPegaProduto("PR_Linha") & "," & RSPegaProduto("PR_Secao") & "," & WCSPROD & ", " _
                    & "" & ConverteVirgula(Format(RSPegaProduto("PR_PrecoVenda1"), "0.00")) & "," & ConverteVirgula(Format(0, "0.00")) & "," & ConverteVirgula(RSPegaProduto("PR_ICMPDV")) & ", " _
                    & "'" & RSPegaProduto("PR_CodigoBarra") & "'," & txtNF.Text & ", '" & txtSerie.Text & "'," & txtCliente.Text & ", " _
                    & "'E'," & Numeros(txtVendedor.Text) & ",'" & txtLjOrigem.Text & "',24,'A'," & ConverteVirgula(0) & "," & ConverteVirgula(0) & ",'0','0','0' , '0')"
                
                rdoCnLojaBach.Execute (SQL)
                If Err.Number = 0 Then
                   CommitTrans
                Else
                   Rollback
                   MsgBox "Não foi possível efetuar a gravação dos itens", vbCritical, "Atenção!"
                End If
            
            Next i
      
   End Select
   
    SQL = ""
    SQL = "Update NFItens Set TipoNota = '" & WTipoNota & "', TipoMovimentacao = " & wTipoMovimentacao & " " _
        & "Where NF = " & txtNF.Text & " and Serie = '" & txtSerie.Text & "'"
    rdoCnLojaBach.Execute (SQL)
      
   rdoProduto.Close
   
End Sub

Sub Alteracao()

    Dim VerificaICMS As Variant
    Dim i As Long

   BeginTrans
   
   If mskDataEmi <> Wdata Then
      MsgBox "Você não pode alterar esta nota. Nota Fiscal de outro dia.", vbInformation, "Operações Especiais"
      Exit Sub
    End If
   Select Case txtCFO.Text
      Case 522
            rdoCnLojaBach.Execute "Update NfCapa set CodOper=" & txtCFO.Text & ", NF=" & txtNF.Text & ", Serie='" & txtSerie.Text & "', LojaOrigem='" & txtLjOrigem.Text & "', " _
                          & "LojaT='" & txtLjDestino.Text & "', DataEmi='" & Format(mskDataEmi.Text, "mm/dd/yyyy") & "', CGCCLi='" & txtCGC.Text & "', ValorMercadoria=" & ConverteVirgula(Format(txtSubTotal.Text, "0.00")) & ", " _
                          & "TotalNota = " & ConverteVirgula(txtTotal.Text) & ", " _
                          & "TipoNota='T', SituacaoEnvio='A', Pesobr=" & ConverteVirgula(SomaPeso) & ", PesoLq=" & ConverteVirgula(SomaPeso) & ", NumeroPed=0, Cliente=0, Hora= DateTime  where LojaOrigem = '" & Trim(txtLjOrigem) & "' and NF= " & Val(txtNF.Text) & " and Serie='" & Trim(txtSerie.Text) & "'"
                            
            WPERDESC = Val(txtDesconto / (txtSubTotal - txtDesconto)) * 100
            
            For i = 1 To grdItens.Rows - 1
               Set rdoProduto = Conexao.OpenResultset("Select PR_PrecoVenda1 from Produto where PR_Referencia = '" & grdItens.TextMatrix(i, 0) & "'")

               wVlDescRat = grdItens.TextMatrix(i, 3) * WPERDESC / 100
               
               rdoCnLojaBach.Execute "Update NfItens set NF = " & txtNF.Text & ", Serie='" & txtSerie.Text & "', LojaOrigem='" & txtLjOrigem.Text & "', DataEmi='" & Format(mskDataEmi.Text, "mm/dd/yyyy") & "', Item=" & i & ", " _
                              & "Referencia = '" & grdItens.TextMatrix(i, 0) & "', Qtde=" & grdItens.TextMatrix(i, 2) & ", VlUnit=" & ConverteVirgula(grdItens.TextMatrix(i, 3)) & ", VlUnit2=" & ConverteVirgula(grdItens.TextMatrix(i, 4) - wVlDescRat) & ", " _
                              & "PLista= " & ConverteVirgula(Format(rdoProduto("PR_PrecoVenda1"), "0.00")) & ", " _
                              & "TipoNota='T', SituacaoEnvio='A' Where LojaOrigem= '" & Trim(txtLjOrigem) & "' and NF= " & Val(txtNF.Text) & " and Serie= '" & Trim(txtSerie.Text) & "' and Referencia='" & grdItens.TextMatrix(i, 0) & "'"
            Next i
            
      Case 5102, 612, 519
         If Val(cmbCondPagto.Text) = 85 Then
            rdoCnLojaBach.Execute "Update NfCapa set CodOper=" & txtCFO.Text & " , NF= " & txtNF.Text & ", Serie='" & txtSerie.Text & "', LojaOrigem='" & txtLjOrigem.Text & "', " _
                           & "Vendedor=" & Val(txtVendedor.Text) & " ,DataEmi='" & Format(mskDataEmi.Text, "mm/dd/yyyy") & "', CondPag=" & Left(cmbCondPagto.Text, 2) & ", " _
                           & "DataPag= '" & Format(mskDataPagto, "mm/dd/yyyy") & "', PgEntra=" & ConverteVirgula(txtPagtoEnt.Text) & ", Cliente = " & txtCliente.Text & ", """ _
                           & "CGCCli='" & txtCGC.Text & "',InscriCli='" & txtInscEst.Text & "', NomCli='" & txtRazao.Text & "', EndCli='" & txtEndereco.Text & "', " _
                           & "BairroCli='" & txtBairro.Text & "', MunicipioCli='" & txtMunicipio.Text & "', UFCliente='" & cmbUFCli.Text & "', CepCli='" & txtCEP.Text & "'," _
                           & "ValorMercadoria=" & ConverteVirgula(txtSubTotal.Text) & ", Desconto=" & ConverteVirgula(MatItens(Linhas).Desconto) & ", " _
                           & "ValFrete= " & ConverteVirgula(txtFrete.Text) & ", FreteCobr=" & ConverteVirgula(txtFrete.Text) & ", TotalNota=" & ConverteVirgula(txtTotal.Text) & ", TipoNota='V', " _
                           & "SituacaoEnvio='A', PesoLq=" & ConverteVirgula(SomaPeso) & ", PesoBr=" & ConverteVirgula(SomaPeso) & ", Hora=DateTime " _
                           & "where NF=" & txtNF.Text & " and Serie= '" & txtSerie.Text & "' and LojaOrigem='" & txtLjOrigem.Text & "'"
         Else
            rdoCnLojaBach.Execute "Update NfCapa set CodOper=" & txtCFO.Text & ", NF=" & txtNF.Text & ", Serie='" & txtSerie.Text & "', LojaOrigem='" & txtLjOrigem & "', Vendedor=" & Val(txtVendedor.Text) & " , " _
                           & "DataEmi='" & Format(mskDataEmi.Text, "mm/dd/yyyy") & "', CondPag=" & Val(cmbCondPagto.Text) & ", PgEntra=" & ConverteVirgula(Format(txtPagtoEnt.Text, "0.00")) & ", Cliente=" & Val(txtCliente.Text) & ", " _
                           & "CGCCli='" & txtCGC.Text & "', InscriCli='" & txtInscEst.Text & "', NomCli='" & txtRazao.Text & "', EndCli='" & txtEndereco.Text & "', BairroCli='" & txtBairro.Text & "', " _
                           & "MunicipioCli= '" & txtMunicipio.Text & "', UFCliente='" & cmbUFCli.Text & "', CepCli='" & txtCEP.Text & "', ValorMercadoria=" & ConverteVirgula(Format(txtSubTotal.Text, "0.00")) & ", " _
                           & "Desconto=" & ConverteVirgula(Format(MatItens(Linhas).Desconto, "0.00")) & ", " _
                           & "ValorFrete=" & ConverteVirgula(Format(txtFrete.Text, "0.00")) & ",  " _
                           & "FreteCobr=" & ConverteVirgula(Format(txtFrete.Text, "0.00")) & ", TotalNota=" & ConverteVirgula(Format(txtTotal.Text, "0.00")) & ", TipoNota='V', VC_SituacaoEnvio='A', " _
                           & "PesoLq=" & ConverteVirgula(Format(SomaPeso, "0.00")) & ", PesoBr=" & ConverteVirgula(Format(SomaPeso, "0.00")) & ", Hora=DateTime where NF=" & txtNF.Text & " and " _
                           & "Serie='" & txtSerie.Text & "' and LojaOrigem='" & txtLjOrigem.Text & "'"
                           
         End If
         
         For i = 1 To grdItens.Rows - 1
            Set rdoProduto = Conexao.OpenResultset("Select PR_PrecoVenda1 from Produto where PR_Referencia = '" & grdItens.TextMatrix(i, 0) & "'")

            rdoCnLojaBach.Execute "Update NfItens set NF=" & txtNF.Text & ", Serie='" & txtSerie.Text & "', LojaOrigem='" & txtLjOrigem.Text & "', DataEmi='" & Format(mskDataEmi.Text, "mm/dd/yyyy") & "', Item=" & i & ",  " _
                           & "Referencia='" & grdItens.TextMatrix(i, 0) & "', Qtde=" & grdItens.TextMatrix(i, 2) & ", VlUnit=" & ConverteVirgula(Format(grdItens.TextMatrix(i, 3), "0.00")) & ", Desconto=" & ConverteVirgula(Format(MatItens(i).Desconto, "0.00")) & ", " _
                           & "PLista=" & ConverteVirgula(Format(rdoProduto("PR_PrecoVenda1"), "0.00")) & ", VlUnit2= " & ConverteVirgula(Format(grdItens.TextMatrix(i, 4), "0.00") - wVlDescRat) & ", " _
                           & "PesoBr=" & ConverteVirgula(Format(Peso, "0.00")) & ", PesoLq=" & ConverteVirgula(Format(Peso, "0.00")) & ", TipoNota='V', SituacaoEnvio='A' where NF= " & txtNF.Text & " and Serie= '" & txtSerie.Text & "' and " _
                           & "Referencia='" & grdItens.TextMatrix(i, 0) & "' and LojaOrigem='" & txtLjOrigem.Text & "'"
         Next i
         
      Case 599, 699, 199, 299
         rdoCnLojaBach.Execute "Update NfCapa set CodOper=" & txtCFO.Text & ", NF=" & txtNF.Text & ", Serie='" & txtSerie.Text & "', LojaOrigem='" & txtLjOrigem & "', DataEmi='" & Format(mskDataEmi.Text, "mm/dd/yyyy") & "', " _
                        & "Cliente=" & txtCliente.Text & ", CGCCli='" & txtCGC.Text & "', InscriCli='" & txtInscEst.Text & "', NomCli='" & txtRazao.Text & "', EndCli='" & txtEndereco.Text & "', BairroCli='" & txtBairro.Text & "', MunicipioCli='" & txtMunicipio.Text & "', " _
                        & "UFCliente='" & cmbUFCli.Text & "', CepCli='" & txtCEP.Text & "', ValorMercadoria=" & ConverteVirgula(Format(txtSubTotal.Text, "0.00")) & ", " _
                        & "TotalNota=" & ConverteVirgula(Format(txtTotal.Text, "0.00")) & ", TipoNota='E', SituacaoEnvio='A', " _
                        & "PesoLq=" & ConverteVirgula(Format(SomaPeso, "0.00")) & ", PesoBr=" & ConverteVirgula(Format(SomaPeso, "0.00")) & ", Hora=DateTime where NF=" & txtNF.Text & " and Serie='" & txtSerie.Text & "' and LojaOrigem= '" & txtLjOrigem.Text & "'"
                        
        
         For i = 1 To grdItens.Rows - 1
            Set rdoProduto = Conexao.OpenResultset("Select PR_PrecoVenda1 from Produto where PR_Referencia = '" & grdItens.TextMatrix(i, 0) & "'")
 
            rdoCnLojaBach.Execute "Update NfItens set NF=" & txtNF.Text & ", Serie='" & txtSerie.Text & "', LojaOrigem='" & txtLjOrigem.Text & "', DataEmi='" & Format(mskDataEmi.Text, "mm/dd/yyyy") & "', Item=" & i & ", Referencia='" & grdItens.TextMatrix(i, 0) & "', Qtde=" & grdItens.TextMatrix(i, 2) & "," _
                           & "VlUnit=" & ConverteVirgula(Format(grdItens.TextMatrix(i, 3), "0.00")) & ", PLista=" & ConverteVirgula(Format(rdoProduto("PR_PrecoVenda1"), "0.00")) & ", VlUnit2=" & ConverteVirgula(Format(grdItens.TextMatrix(i, 4), "0.00") - wVlDescRat) & ", " _
                           & "PesoBr=" & ConverteVirgula(Format(Peso, "0.00")) & ", PesoLiq=" & ConverteVirgula(Format(Peso, "0.00")) & ", TipoNota='E', SituacaoEnvio='A' " _
                           & "where NF= " & txtNF.Text & " and Serie= '" & txtSerie.Text & "' and Referencia='" & grdItens.TextMatrix(i, 0) & "' and LojaOrigem='" & txtLjOrigem.Text & "'"
                          
         Next i
      
      
      Case 132, 232
         rdoCnLojaBach.Execute "Update NfCapa set CodOper=" & txtCFO.Text & ", NF=" & txtNF.Text & ", Serie='" & txtSerie.Text & "', LojaOrigem='" & txtLjOrigem & "', Vendedor=" & Val(txtVendedor.Text) & ", DataEmi='" & Format(mskDataEmi.Text, "mm/dd/yyyy") & "', Cliente=" & txtCliente.Text & ", " _
                        & "CGCCli='" & txtCGC.Text & "', InscriCli='" & txtInscEst.Text & "', NomCli='" & txtRazao.Text & "', EndCli='" & txtEndereco.Text & "', BairroCli='" & txtBairro.Text & "', MunicipioCli='" & txtMunicipio.Text & "', UFCliente='" & cmbUFCli.Text & "', CepCli='" & txtCEP.Text & "', " _
                        & "ValorMercadoria=" & ConverteVirgula(Format(txtSubTotal.Text, "0.00")) & ", " _
                        & "TotalNota=" & ConverteVirgula(Format(txtTotal.Text, "0.00")) & ", TipoNota='E', VC_SituacaoEnvio='A', PesoLq=" & ConverteVirgula(Format(SomaPeso, "0.00")) & ", PesoBr=" & ConverteVirgula(Format(SomaPeso, "0.00")) & ", " _
                        & " Hora=DateTime where NF= " & txtNF.Text & " and Serie='" & txtSerie.Text & "' and LojaOrigem='" & txtLjOrigem.Text & "' "
         
         For i = 1 To grdItens.Rows - 1
            Set rdoProduto = Conexao.OpenResultset("Select PR_PrecoVenda1 from Produto where PR_Referencia = '" & grdItens.TextMatrix(i, 0) & "'")
            
            rdoCnLojaBach.Execute "Update NfItens set NF=" & txtNF.Text & ", Serie='" & txtSerie.Text & "', LojaOrigem='" & txtLjOrigem.Text & "', DataEmi='" & Format(mskDataEmi.Text, "mm/dd/yyyy") & "', Item=I, Referencia='" & grdItens.TextMatrix(i, 0) & "', Qtde=" & grdItens.TextMatrix(i, 2) & ", " _
                           & "VlUnit=" & ConverteVirgula(Format(grdItens.TextMatrix(i, 3), "0.00")) & ", PLista=" & ConverteVirgula(Format(rdoProduto("PR_PrecoVenda1"), "0.00")) & ", VlUnit2=" & ConverteVirgula(Format(grdItens.TextMatrix(i, 4), "0.00") - wVlDescRat) & ", " _
                           & "PesoBr=" & ConverteVirgula(Format(Peso, "0.00")) & ", PesoLq=" & ConverteVirgula(Format(Peso, "0.00")) & ", TipoNota='E', SituacaoEnvio='A' where " _
                           & "NF= " & txtNF.Text & " and Serie= '" & txtSerie.Text & "' and LojaOrigem= '" & txtLjOrigem.Text & "'and Referencia='" & grdItens.TextMatrix(i, 0) & "'"
         Next i
      
      
      Case 532, 632
         rdoCnLojaBach.Execute "Update NfCapa set CodOper=" & txtCFO.Text & ", NF=" & txtNF.Text & ", Serie='" & txtSerie.Text & "', LojaOrigem='" & txtLjOrigem & "', DataEmi='" & Format(mskDataEmi.Text, "mm/dd/yyyy") & "', Cliente=" & txtCliente.Text & ", CGCCli='" & txtCGC.Text & "', InscriCli='" & txtInscEst.Text & "', " _
                       & "NomCli='" & txtRazao.Text & "', EndCli='" & txtEndereco.Text & "', BairroCli='" & txtBairro.Text & "', MunicipioCli='" & txtMunicipio.Text & "', UFCliente='" & cmbUFCli.Text & "', CepCli='" & txtCEP.Text & "', ValorMercadoria=" & ConverteVirgula(Format(txtSubTotal.Text, "0.00")) & ", " _
                       & "TotalNota=" & ConverteVirgula(Format(txtTotal.Text, "0.00")) & ", TipoNota='S', " _
                       & "SituacaoEnvio='A', PesoLq=" & ConverteVirgula(Format(SomaPeso, "0.00")) & ", PesoBr=" & ConverteVirgula(Format(SomaPeso, "0.00")) & ", Hora=DateTime where NF=" & txtNF.Text & " and Serie='" & txtSerie.Text & "' and LojaOrigem='" & txtLjOrigem.Text & "'"
                       
         
         For i = 1 To grdItens.Rows - 1
            Set rdoProduto = Conexao.OpenResultset("Select PR_PrecoVenda1 from Produto where PR_Referencia = '" & grdItens.TextMatrix(i, 0) & "'")

            rdoCnLojaBach.Execute "Update NfItens set NF=" & txtNF.Text & ", Serie='" & txtSerie.Text & "', LojaOrigem='" & txtLjOrigem.Text & "', DataEmi='" & Format(mskDataEmi.Text, "mm/dd/yyyy") & "', Item=" & i & ", Referencia='" & grdItens.TextMatrix(i, 0) & "', Qtde=" & grdItens.TextMatrix(i, 2) & ", " _
                           & "VlUnit=" & ConverteVirgula(Format(grdItens.TextMatrix(i, 3), "0.00")) & ", PLista=" & ConverteVirgula(Format(rdoProduto("PR_PrecoVenda1"), "0.00")) & ", VlUnit2=" & ConverteVirgula(Format(grdItens.TextMatrix(i, 4), "0.00") - wVlDescRat) & ", " _
                           & "PesoBr=" & ConverteVirgula(Format(Peso, "0.00")) & ", PesoLq=" & ConverteVirgula(Format(Peso, "0.00")) & ", TipoNota='S', SituacaoEnvio='A' " _
                           & "where NF= " & txtNF.Text & " and Serie= '" & txtSerie.Text & "' and LojaOrigem='" & txtLjOrigem.Text & "' and Referencia='" & grdItens.TextMatrix(i, 0) & "'"
         
         Next i
      
   End Select
   
   
   
   If Err.Number = 0 Then
      GravacaoOK = True
      CommitTrans
   Else
      GravacaoOK = False
      Rollback
      MsgBox "Não foi possível efetuar a alteração", vbInformation, "Atenção!"
   End If
   
   AlteraDados = False
   LimpaTela

End Sub

Sub LimpaTela()

   SubTotal = 0
   txtCFO.Text = ""
   cmbNaturezaOperacao.Visible = True
   cmbNaturezaOperacao.ListIndex = -1
   txtNF.Text = ""
   txtSerie.Text = ""
   txtLjOrigem.Text = ""
   txtLjDestino.Text = ""
   txtVendedor.Text = ""
   txtOutroVendedor.Text = ""
   mskDataEmi.Text = Format(Wdata, "dd/mm/yyyy")
   cmbCondPagto.ListIndex = -1
   mskDataPagto.Text = "__/__/____"
   txtPagtoEnt.Text = ""
   txtCliente.Text = ""
   txtCGC.Text = ""
   txtInscEst.Text = ""
   txtRazao.Text = ""
   txtEndereco.Text = ""
   txtBairro.Text = ""
   txtMunicipio.Text = ""
   cmbUFCli.ListIndex = -1
   txtCEP.Text = ""
   txtFoneCli.Text = ""
   txtreferencia.Text = ""
   txtdescricao.Text = ""
   txtQuant.Text = ""
   txtPrecoUnit.Text = ""
   txtTotalItem.Text = ""
   txtSubTotal.Text = 0
   txtDesconto.Text = ""
   txtValorIPI.Text = ""
   txtAliqIPI.Text = ""
   txtFrete.Text = ""
   txtTotal.Text = ""
   grdItens.Rows = 2
   grdItens.AddItem ""
   grdItens.RemoveItem (1)
   txtQuant.Enabled = False
   txtPrecoUnit.Enabled = False
   txtNF.SetFocus
   
End Sub

Private Sub txtVendedor_LostFocus()

   If GetAsyncKeyState(vbKeyTab) <> 0 Then
      If txtVendedor.Text <> "" Then
         Set rdoVendedor = Conexao.OpenResultset("Select VE_CodigoVendedor, VE_Nome from Vendedor where VE_CodigoVendedor= " & Numeros(txtVendedor.Text) & " and VE_Loja = '" & txtLjOrigem.Text & "'")
         
         If rdoVendedor.EOF Then
            MsgBox "Vendedor não encontrado.", vbInformation, "Atenção"
            txtVendedor.SelStart = 0
            txtVendedor.SelLength = Len(txtVendedor.Text)
            txtVendedor.SetFocus
         Else
            txtVendedor.Text = rdoVendedor("VE_CodigoVendedor") & " - " & rdoVendedor("VE_Nome")
         End If
         
         rdoVendedor.Close
         
         If txtLjOrigem.Text = txtLjDestino.Text Then
             cmbCondPagto.SetFocus
         Else
             txtOutroVendedor.SetFocus
         End If
         
      End If
   End If

End Sub

Sub PreencheComboUF(ByRef cmbEstado As ComboBox)
   cmbEstado.AddItem "AC"
   cmbEstado.AddItem "AL"
   cmbEstado.AddItem "AM"
   cmbEstado.AddItem "AP"
   cmbEstado.AddItem "BA"
   cmbEstado.AddItem "CE"
   cmbEstado.AddItem "DF"
   cmbEstado.AddItem "ES"
   cmbEstado.AddItem "GO"
   cmbEstado.AddItem "MG"
   cmbEstado.AddItem "MS"
   cmbEstado.AddItem "MT"
   cmbEstado.AddItem "PA"
   cmbEstado.AddItem "PB"
   cmbEstado.AddItem "PE"
   cmbEstado.AddItem "PI"
   cmbEstado.AddItem "PR"
   cmbEstado.AddItem "SC"
   cmbEstado.AddItem "SE"
   cmbEstado.AddItem "SP"
   cmbEstado.AddItem "RJ"
   cmbEstado.AddItem "RN"
   cmbEstado.AddItem "RO"
   cmbEstado.AddItem "RR"
   cmbEstado.AddItem "RS"
   cmbEstado.AddItem "TO"
End Sub

Sub MontaComboNatureza(ByVal CFO As Long, ByRef cmbNaturezaOperacao As ComboBox)
    
    Dim Ponteiro As Long
    Dim Maximo As Long
    
    cmbNaturezaOperacao.Clear
    Ponteiro = 0
    Maximo = UBound(matNatureza)
    If CFO > 0 Then
        Do While CFO <> matNatureza(Ponteiro).CFO
            Ponteiro = Ponteiro + 1
            If Ponteiro > Maximo Then
                Exit Sub
            End If
        Loop
        
        Do While CFO = matNatureza(Ponteiro).CFO
            cmbNaturezaOperacao.AddItem matNatureza(Ponteiro).Descricao
            Ponteiro = Ponteiro + 1
            If Ponteiro > Maximo Then
                Exit Sub
            End If
        Loop
    End If
    
End Sub

Sub CarregaNaturezaOperacao()
    
    Dim rdoCombos As rdoResultset
    Dim Ponteiro As Long
    
    Set rdoCombos = Conexao.OpenResultset("Select CN_CodigoOperacaoNovo, CN_DescricaoOperacao from CodigoOperacaoNovo order by CN_CodigoOperacaoNovo", Options:=rdExecDirect)
    
    ReDim matNatureza(0) As Natureza
    
    Ponteiro = 0
    Do While Not rdoCombos.EOF
        ReDim Preserve matNatureza(Ponteiro) As Natureza
        matNatureza(Ponteiro).CFO = rdoCombos("CN_CodigoOperacaoNovo")
        matNatureza(Ponteiro).Descricao = rdoCombos("CN_DescricaoOperacao")
        Ponteiro = Ponteiro + 1
        rdoCombos.MoveNext
    Loop
    
    rdoCombos.Close
    
End Sub
Function EncontraTexto(ByRef Combo As ComboBox) As Boolean

    Dim Indice As Integer
    Dim Maximo As Integer
    
    Maximo = Combo.ListCount - 1
    
    If Len(Trim$(Combo.Text)) > 0 Then
        If Val(Combo.Text) <> 0 Then
            For Indice = 0 To Maximo Step 1
                If Val(Combo.Text) = Combo.ItemData(Indice) Then
                    Combo.ListIndex = Indice
                    EncontraTexto = True
                    Exit Function
                End If
            Next Indice
        Else
            For Indice = 0 To Maximo Step 1
                If UCase$(Combo.Text) = Left$(Trim$(Mid$(UCase$(Combo.List(Indice)), (InStr(Combo.List(Indice), "-") + 1))), Len(Combo.Text)) Then
                    Combo.ListIndex = Indice
                    EncontraTexto = True
                    Exit Function
                End If
            Next Indice
        End If
    End If
    
    EncontraTexto = False

End Function

Public Function ExtraiSeqPedidoDbfOperEspecial()

     Dim WNovoSeqPed As Long
     
     
        WnumeroPedidoDbf = 0
        WNovoSeqPed = 0
        WnumeroPed = 0
        
'        Set DBFBanco = Workspaces(0).OpenDatabase(WbancoDbf, False, False, "DBase IV")
'
'        Set RsDadosDbf = DBFBanco.OpenResultset("Select * from controle.dbf ")
        
        WnumeroPedidoDbf = RsDadosDbf("NumPed") + 1
        WNovoSeqPed = WnumeroPedidoDbf
        WnumeroPed = WnumeroPedidoDbf
        
        BeginTrans
        
        SQL = "update controle.dbf set NumPed= " & WNovoSeqPed & ""
        DBFBanco.Execute (SQL)
        
        CommitTrans
     
        DBFBanco.Close
        

End Function

Sub vertecla(ByRef Tecla As Integer)
    
    If Not IsNumeric(Chr(Tecla)) And Tecla <> vbKeyBack Then
        Tecla = 0
    End If

End Sub

Sub VerteclaVirgula(ByRef Controle As Control, ByRef Tecla As Integer)

    If Controle.SelStart = 0 And Controle.SelLength = Len(Controle.Text) Then
        Controle.Text = ""
    End If

    If Chr(Tecla) = "," Or Chr(Tecla) = "." Then
        If InStr(Controle.Text, ",") <> 0 Or InStr(Controle.Text, ".") <> 0 Then
            Tecla = 0
        Else
            Tecla = Asc(",")
        End If
    ElseIf Not IsNumeric(Chr(Tecla)) And Tecla <> vbKeyBack Then
        Tecla = 0
    End If

End Sub

Function Numeros(ByVal Texto As Variant) As String

    Dim Maximo As Integer
    Dim Char As Integer
    Dim CharLido As String * 1
    Dim Retorno As String
    
    If IsNull(Texto) Then
        Numeros = ""
        
        Exit Function
    End If
    
    Maximo = Len(Texto)
    
    Retorno = ""
    For Char = 1 To Maximo Step 1
        CharLido = Mid(Texto, Char, 1)
        If IsNumeric(CharLido) Then
            Retorno = Retorno & CharLido
        End If
    Next Char
    
    Texto = Retorno
    
    Numeros = Texto

End Function

