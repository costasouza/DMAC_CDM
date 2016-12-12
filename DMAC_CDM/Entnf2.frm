VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmEntradaNotaFiscal2 
   BackColor       =   &H8000000A&
   Caption         =   "Entrada Nota Fiscal Fornecedores"
   ClientHeight    =   7605
   ClientLeft      =   1575
   ClientTop       =   2265
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7605
   ScaleWidth      =   12000
   Begin VB.PictureBox PtbSistema 
      Height          =   555
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   11745
      TabIndex        =   72
      Top             =   6990
      Width           =   11805
      Begin VB.CommandButton cmdVencimentos 
         Caption         =   "Vencimentos"
         Enabled         =   0   'False
         Height          =   495
         Left            =   4545
         TabIndex        =   78
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton cmdEntraNota 
         Caption         =   "Limpa"
         Height          =   495
         Index           =   4
         Left            =   5745
         TabIndex        =   77
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton cmdEntraNota 
         Caption         =   "Encerra Nota"
         Height          =   495
         Index           =   3
         Left            =   6945
         TabIndex        =   76
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton cmdEntraNota 
         Caption         =   "Exclui Nota"
         Height          =   495
         Index           =   2
         Left            =   8145
         TabIndex        =   75
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton cmdEntraNota 
         Caption         =   "Grava Capa"
         Height          =   495
         Index           =   1
         Left            =   9345
         TabIndex        =   74
         Top             =   0
         Width           =   1200
      End
      Begin VB.CommandButton cmdEntraNota 
         Caption         =   "Retorna"
         Height          =   495
         Index           =   0
         Left            =   10545
         TabIndex        =   73
         Top             =   0
         Width           =   1200
      End
   End
   Begin Threed.SSPanel pnlObsPedido 
      Height          =   375
      Left            =   2070
      TabIndex        =   68
      Top             =   2985
      Visible         =   0   'False
      Width           =   4050
      _Version        =   65536
      _ExtentX        =   7144
      _ExtentY        =   661
      _StockProps     =   15
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   0
      BorderWidth     =   0
      BevelOuter      =   0
      Outline         =   -1  'True
      Alignment       =   0
      Autosize        =   2
   End
   Begin Threed.SSPanel pnlItens 
      Height          =   645
      Left            =   6000
      TabIndex        =   58
      Top             =   6285
      Width           =   5940
      _Version        =   65536
      _ExtentX        =   10477
      _ExtentY        =   1138
      _StockProps     =   15
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Begin VB.TextBox txtPercentualDesconto 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   4995
         MaxLength       =   7
         TabIndex        =   30
         Top             =   240
         Width           =   825
      End
      Begin VB.TextBox txtAliquotaIPI 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   4320
         MaxLength       =   7
         TabIndex        =   29
         Top             =   240
         Width           =   630
      End
      Begin VB.TextBox txtPrecoUnitario 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   3015
         MaxLength       =   17
         TabIndex        =   28
         Top             =   240
         Width           =   1260
      End
      Begin VB.TextBox txtQuantidade 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   2310
         MaxLength       =   5
         TabIndex        =   27
         Top             =   240
         Width           =   660
      End
      Begin VB.TextBox txtNumeroPedido 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   1335
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   240
         Width           =   930
      End
      Begin VB.TextBox txtReferencia 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   90
         MaxLength       =   7
         TabIndex        =   25
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% Desc."
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5025
         TabIndex        =   64
         Top             =   30
         Width           =   585
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% IPI"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4350
         TabIndex        =   63
         Top             =   30
         Width           =   360
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preço Unitário"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3030
         TabIndex        =   62
         Top             =   30
         Width           =   1005
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde."
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2325
         TabIndex        =   61
         Top             =   30
         Width           =   390
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pedido"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1335
         TabIndex        =   60
         Top             =   30
         Width           =   495
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Referência"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   105
         TabIndex        =   59
         Top             =   30
         Width           =   780
      End
   End
   Begin VB.TextBox txtPesquisaDescricao 
      BackColor       =   &H00C0C0C0&
      Height          =   315
      Left            =   1680
      MaxLength       =   38
      TabIndex        =   32
      Top             =   6525
      Width           =   4245
   End
   Begin VB.TextBox txtPesquisaPedido 
      BackColor       =   &H00C0C0C0&
      Height          =   315
      Left            =   795
      MaxLength       =   6
      TabIndex        =   31
      Top             =   6525
      Width           =   840
   End
   Begin MSFlexGridLib.MSFlexGrid grdItens 
      Height          =   3495
      Left            =   60
      TabIndex        =   24
      Top             =   2640
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   6165
      _Version        =   393216
      Cols            =   17
      FixedCols       =   0
      BackColor       =   16777215
      ForeColorFixed  =   192
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      Enabled         =   0   'False
      FocusRect       =   0
      SelectionMode   =   1
   End
   Begin VB.Frame Frame1 
      Height          =   2490
      Left            =   75
      TabIndex        =   33
      Top             =   0
      Width           =   9390
      Begin VB.TextBox txtaliquotaicms 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   2790
         TabIndex        =   12
         Top             =   1530
         Width           =   1260
      End
      Begin VB.ComboBox cmbNaturezaOperacao 
         BackColor       =   &H80000000&
         Height          =   315
         Left            =   4830
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   885
         Width           =   4395
      End
      Begin VB.TextBox txtValorMercadorias 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   1530
         Width           =   1335
      End
      Begin VB.TextBox txtEmbalagem 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   2070
         Width           =   1335
      End
      Begin VB.TextBox txtOutros 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   4095
         TabIndex        =   20
         Top             =   2070
         Width           =   1260
      End
      Begin VB.TextBox txtDespesas 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1500
         TabIndex        =   18
         Top             =   2070
         Width           =   1260
      End
      Begin VB.TextBox txtJuros 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   2790
         TabIndex        =   19
         Top             =   2070
         Width           =   1260
      End
      Begin VB.TextBox txtValorIPI 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   8010
         TabIndex        =   16
         Top             =   1530
         Width           =   1260
      End
      Begin VB.TextBox txtFrete 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   6705
         TabIndex        =   15
         Top             =   1530
         Width           =   1260
      End
      Begin VB.TextBox txtBaseICMS 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1500
         TabIndex        =   11
         Top             =   1530
         Width           =   1260
      End
      Begin VB.TextBox txtValorICMS 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   4110
         TabIndex        =   13
         Top             =   1530
         Width           =   1260
      End
      Begin VB.TextBox txtValorICMSSubsTrib 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   5400
         TabIndex        =   14
         Top             =   1530
         Width           =   1260
      End
      Begin VB.TextBox txtValorTotalNota 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   5400
         TabIndex        =   21
         Top             =   2070
         Width           =   1260
      End
      Begin VB.TextBox txtTotalCalculado 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   6705
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   2070
         Width           =   1260
      End
      Begin VB.TextBox txtBateNota 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   8010
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2070
         Width           =   1260
      End
      Begin MSMask.MaskEdBox mskDataEmissao 
         Height          =   315
         Left            =   1353
         TabIndex        =   6
         Top             =   885
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483648
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cmbLocal 
         BackColor       =   &H80000004&
         Height          =   315
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   345
         Width           =   1230
      End
      Begin VB.TextBox txtNomeFantasia 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   3951
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   345
         Width           =   3420
      End
      Begin VB.TextBox txtSerie 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   8595
         MaxLength       =   2
         TabIndex        =   5
         Top             =   345
         Width           =   585
      End
      Begin VB.TextBox txtNotaFiscal 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   7494
         TabIndex        =   4
         Top             =   345
         Width           =   975
      End
      Begin VB.TextBox txtCodigoOperacao 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   3819
         MaxLength       =   4
         TabIndex        =   8
         Top             =   885
         Width           =   900
      End
      Begin VB.TextBox txtCGC 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   1503
         MaxLength       =   15
         TabIndex        =   2
         Top             =   345
         Width           =   2325
      End
      Begin VB.TextBox txtDataEntrada 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   135
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   885
         Width           =   1110
      End
      Begin MSMask.MaskEdBox mskDataRecebimento 
         Height          =   315
         Left            =   2586
         TabIndex        =   7
         Top             =   885
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483648
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblnaturezaoperacao 
         Caption         =   "lblnaturezaoperacao"
         Height          =   180
         Left            =   6630
         TabIndex        =   67
         Top             =   690
         Visible         =   0   'False
         Width           =   2430
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Aliq. ICMS"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2850
         TabIndex        =   69
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "Data Receb."
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2595
         TabIndex        =   66
         Top             =   675
         Width           =   915
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "Natureza de Operação"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4860
         TabIndex        =   65
         Top             =   675
         Width           =   1620
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000009&
         X1              =   20
         X2              =   9370
         Y1              =   1290
         Y2              =   1290
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000C&
         X1              =   20
         X2              =   9370
         Y1              =   1275
         Y2              =   1275
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Valor Mercadorias"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   150
         TabIndex        =   55
         Top             =   1320
         Width           =   1275
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Embalagem"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   150
         TabIndex        =   54
         Top             =   1860
         Width           =   825
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Outros"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   4110
         TabIndex        =   53
         Top             =   1860
         Width           =   465
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Despesas"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1530
         TabIndex        =   52
         Top             =   1860
         Width           =   720
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Valor Juros"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   2820
         TabIndex        =   51
         Top             =   1860
         Width           =   780
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "IPI"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   8055
         TabIndex        =   50
         Top             =   1320
         Width           =   195
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Frete"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   6735
         TabIndex        =   49
         Top             =   1320
         Width           =   360
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Base ICMS"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1530
         TabIndex        =   48
         Top             =   1320
         Width           =   795
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "ICMS"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4140
         TabIndex        =   47
         Top             =   1320
         Width           =   390
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "ICMS Subst."
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5445
         TabIndex        =   46
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Total Nota Fiscal"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5400
         TabIndex        =   45
         Top             =   1860
         Width           =   1200
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Total Calculado"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   6720
         TabIndex        =   44
         Top             =   1860
         Width           =   1110
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Bate Nota"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   8025
         TabIndex        =   43
         Top             =   1860
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Fornecedor"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3975
         TabIndex        =   41
         Top             =   135
         Width           =   810
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Série"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   8610
         TabIndex        =   40
         Top             =   135
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nota Fiscal"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   7530
         TabIndex        =   39
         Top             =   135
         Width           =   795
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "C.F.O"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3840
         TabIndex        =   38
         Top             =   675
         Width           =   405
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Local"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   165
         TabIndex        =   37
         Top             =   135
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "C.G.C."
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1515
         TabIndex        =   36
         Top             =   135
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Data Emissão"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1350
         TabIndex        =   35
         Top             =   675
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Data Entrada"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   135
         TabIndex        =   34
         Top             =   675
         Width           =   945
      End
   End
   Begin Threed.SSPanel pnlCapaPedido 
      Height          =   2400
      Left            =   9525
      TabIndex        =   80
      Top             =   90
      Width           =   2340
      _Version        =   65536
      _ExtentX        =   4128
      _ExtentY        =   4233
      _StockProps     =   15
      BackColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   0
      BorderWidth     =   1
      BevelInner      =   1
      Begin VB.Label lblEntrega2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "99/9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   15
         TabIndex        =   88
         Top             =   1500
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblCondPagto2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "999/999/999/999/999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   15
         TabIndex        =   87
         Top             =   2070
         Visible         =   0   'False
         Width           =   1950
      End
      Begin VB.Label lblPedido 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pedido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   15
         TabIndex        =   86
         Top             =   45
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label lblComprador 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   15
         TabIndex        =   85
         Top             =   315
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label lblCFOP 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CFOP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   15
         TabIndex        =   84
         Top             =   600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblEntrega 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Semana de Entrega"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   15
         TabIndex        =   83
         Top             =   1260
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Label lblCondPagto 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Condição de Pagamento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   15
         TabIndex        =   82
         Top             =   1785
         Visible         =   0   'False
         Width           =   2085
      End
      Begin VB.Label lblFrete 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Frete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   15
         TabIndex        =   81
         Top             =   870
         Visible         =   0   'False
         Width           =   450
      End
   End
   Begin VB.Label lblControlaTela 
      Caption         =   "ControlaTela"
      Height          =   195
      Left            =   375
      TabIndex        =   79
      Top             =   7710
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label lblNotaOrigem 
      Caption         =   "lblNotaOrigem"
      Height          =   165
      Left            =   2415
      TabIndex        =   71
      Top             =   7710
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Label lblSerieOrigem 
      Caption         =   "0"
      Height          =   165
      Left            =   2415
      TabIndex        =   70
      Top             =   7920
      Width           =   1530
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pesquisa:"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   45
      TabIndex        =   57
      Top             =   6570
      Width           =   690
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1725
      TabIndex        =   56
      Top             =   6315
      Width           =   720
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pedido"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   840
      TabIndex        =   42
      Top             =   6315
      Width           =   495
   End
End
Attribute VB_Name = "frmEntradaNotaFiscal2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim TotalCalculado As Double
Dim TotalNota As Double
Dim Frete As Double
Dim Embalagem As Double
Dim Despesas As Double
Dim Juros As Double
Dim Outros As Double
Dim BateNota As Double
Dim QtdeDistribuicao As Double
Dim adoConsEstoque As New ADODB.Recordset
Dim adoChkFornecedor As New ADODB.Recordset

Dim Fornecedor As String
Dim SomaTotal As Double
Dim SomaCalculado As Double
Dim SomaSub As Double
Dim TotalEntrada As Double
Dim QtdeCalculada As Integer
Dim QtdeCalculo As Double
Dim Residencia As Integer
Dim wFlagAtendimento As String

Dim rsDados As New ADODB.Recordset
Dim rsLoja As New ADODB.Recordset
Dim rsPerc As New ADODB.Recordset
Dim rsSum As New ADODB.Recordset
Dim rsFaltaLoja As New ADODB.Recordset

Dim PercCalculado As Double
Dim TotalGrade As Double
Dim I As Integer
Dim wWhere As String


Dim LinhaCFO As Long
Dim UltimoCFO As Long

Dim Pesquisou As Boolean
Dim NotaJaExiste As Boolean
Dim SoExclusao As Boolean

Dim Data As String

Private Type ObservacaoPedido
    NumeroPedido As Long
    Obs As String
End Type

Dim Observacoes() As ObservacaoPedido

Private WithEvents clsNotaFiscal As Cadastro
Attribute clsNotaFiscal.VB_VarHelpID = -1
Private WithEvents clsGravaItem As Cadastro
Attribute clsGravaItem.VB_VarHelpID = -1
Private WithEvents clsItensNota As ControlaGrid
Attribute clsItensNota.VB_VarHelpID = -1

Dim WCodigoOperacaoVelho As String
Dim WCodigoOperacaoNovo As String


Sub InicializaClasses()
    
    Set clsNotaFiscal = New Cadastro
  
    
    Set clsNotaFiscal.Conexao = ADO_Cn_CD
    
    clsNotaFiscal.NomeFormulario = Me.Name
    clsNotaFiscal.ControlFields = False
    clsNotaFiscal.Controles = "txtDataEntrada; cmbLocal; " _
                        & "txtNomeFantasia; txtNotaFiscal; txtSerie; " _
                        & "mskDataEmissao; mskDataRecebimento; " _
                        & "txtCodigoOperacao; lblNaturezaOperacao; " _
                        & "txtValorMercadorias; " _
                        & "txtBaseICMS; txtaliquotaicms; txtValorICMS; " _
                        & "txtValorICMSSubsTrib; txtFrete; txtValorIPI; " _
                        & "txtEmbalagem; txtDespesas; txtJuros; txtOutros; " _
                        & "txtValorTotalNota; txtTotalCalculado; lblNotaOrigem; lblSerieOrigem"
                
    clsNotaFiscal.Campos = "CC_DataEntrada; CC_Loja; " _
                        & "CC_Fornecedor; CC_NotaFiscal; CC_Serie; " _
                        & "CC_DataEmissao; CC_DataRecebimento; " _
                        & "CC_CodigoOperacao; CC_NaturezaOperacao; " _
                        & "CC_ValorMercadorias; " _
                        & "CC_BaseICMS; CC_AliquotaICMS; CC_ValorICMS; " _
                        & "CC_ValorICMSSubsTrib; CC_Frete; CC_ValorIPI; " _
                        & "CC_Embalagem; CC_Despesas; CC_Juros; CC_Outros; " _
                        & "CC_ValorTotalNota; CC_ValorCalculado; CC_NotaConsignacao; CC_SerieConsignacao"
    
    clsNotaFiscal.TipoFormato = "Data; TotalText; " _
                              & "Valor; Numero; Caractere; " _
                              & "Data; Data; " _
                              & "Numero; Valor; " _
                              & "Decimal; " _
                              & "Decimal; Decimal; Decimal; " _
                              & "Decimal; Decimal; Decimal; " _
                              & "Decimal; Decimal; Decimal; Decimal; " _
                              & "Decimal; Decimal; Numero; Caractere"
    
    clsNotaFiscal.ValorLimpeza = "__/__/____; -1; " _
                              & " ;  ;  ; " _
                              & "__/__/____; __/__/____; " _
                              & " ;  ; " _
                              & "0,00; " _
                              & "0,00; 0,00; 0,00; " _
                              & "0,00; 0,00; 0,00; " _
                              & "0,00; 0,00; 0,00; 0,00; " _
                              & "0,00; 0,00; 0;  "
    
    clsNotaFiscal.FazerVerificacao = "1; 1; " _
                                   & "1; 1; 1; " _
                                   & "1; 1; " _
                                   & "1; 1; " _
                                   & "1; " _
                                   & "1; 1; 1; " _
                                   & "1; 1; 1; " _
                                   & "1; 1; 1; 1; " _
                                   & "1; 1; 0; 0"
                        
    clsNotaFiscal.Limpar
        
    
    Set clsGravaItem = New Cadastro
    Set clsGravaItem.Conexao = ADO_Cn_CD
    
    clsGravaItem.NomeFormulario = Me.Name
    clsGravaItem.ControlFields = False
    clsGravaItem.IncludeExtraFields = True
    clsGravaItem.ExtraFields = "CI_NotaFiscal; CI_Serie; CI_Fornecedor"
    clsGravaItem.Controles = "txtReferencia; txtNumeroPedido; txtQuantidade; " _
                        & "txtPrecoUnitario; txtAliquotaIPI; " _
                        & "txtPercentualDesconto; txtDataEntrada"
    
    clsGravaItem.Campos = "CI_Referencia; CI_NossoPedido; CI_Quantidade; " _
                        & "CI_PrecoUnitario; CI_AliquotaIPI; " _
                        & "CI_PercentualDesconto; CI_DataEntrada"
    
    clsGravaItem.TipoFormato = "Caractere; Numero; Numero; " _
                        & "Moeda; Decimal; " _
                        & "Decimal; Data"
    
    clsGravaItem.ValorLimpeza = " ;  ;  ; " _
                        & " ;  ; " _
                        & " "

    clsGravaItem.FazerVerificacao = "1; 1; 1; " _
                        & "1; 1; " _
                        & "1; 1"
        
    
    Set clsItensNota = New ControlaGrid
    Set clsItensNota.ConexaoGrid = ADO_Cn_CD
    
    clsItensNota.NomeFormulario = Me.Name
    clsItensNota.NomeGrid = "GrdItens"
    clsItensNota.Colunas = 17
    clsItensNota.LinhasVisiveis = 6
    clsItensNota.Campos = "PI_Referencia; PI_NumeroPedido; PR_Descricao; " _
                        & "CI_Quantidade; CI_PrecoUnitario; " _
                        & "CI_AliquotaIPI; CI_PercentualDesconto; Total; " _
                        & "PI_SaldoPedido; CI_Observacao; Comprador; Entrega; " _
                        & "CondPag; Frete; CodigoOperacao; PR_Residencia; PC_LojaReserva"
    clsItensNota.Cabecalho = "^Referência; ^Pedido; <Descrição; ^Qtde; " _
                           & "^Preço Unit.; ^%IPI; ^%Desc; ^Total; ^Saldo; " _
                           & "<Critica; ^Comprador; ^Semana ; ^CondPag; " _
                           & "^Frete; ^CFOP; ^Residencia; ^LojaReserva"
    clsItensNota.Formato = "Caractere; Numero; Caractere; Numero; " _
                         & "Decimal; Decimal; Decimal; Decimal; " _
                         & "Numero; Caractere; Caractere; Numero; " _
                         & "Numero; Decimal; Numero; Numero; Caractere"
    clsItensNota.Tamanho = "850; 570; 3890; 450; 870; 540; 570; 870; 490; " & _
                           "8500; 1; 1; 1; 1; 1; 1; 1"
    clsItensNota.Alinhamento = "Esquerda; Direita; Esquerda; Direita; " _
                             & "Direita; Direita; Direita; Direita; Direita; " _
                             & "Esquerda; Direita; Direita; Direita; Direita; " _
                             & "Direita; Direita; Direita"
                             
    clsItensNota.MontaCabecalho

End Sub

Sub MontaCombos()
    
    Dim adoCombos As New ADODB.Recordset
    
    SQL = "Select LO_Loja from Loja where LO_Loja not in ('CMC','CMCS','CONSO','Alm01') and LO_Situacao = 'A'"
    
    adoCombos.CursorLocation = adUseClient
    adoCombos.Open SQL, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
    
    'Set adoCombos = rdoCnSupBatch.OpenResultset("Select LO_Loja from Loja where LO_Loja not in ('CMC','CMCS','CONSO','Alm01') and LO_Situacao = 'A'", Options:=rdExecDirect)
    PreencheCombo cmbLocal, adoCombos, "LO_Loja", ""
    cmbLocal.ListIndex = GetIndice(Almoxarifado, cmbLocal)

'    Set adoCombos = rdoCnSupBatch.OpenResultset("Select CP_CodigoCondicao, CP_Descricao from CondicaoPagto Where CP_VendaCompra = 'C'", Options:=rdExecDirect)
'    PreencheCombo cmbCondicaoPagto, adoCombos, "CP_CodigoCondicao", "CP_Descricao"
    
    adoCombos.Close
    
End Sub

Sub MostraData()
    
    Dim rdoEntraNota As New ADODB.Recordset
    
    SQL = "Select CV_UltimoDiaMes from ControleFec"
    
    rdoEntraNota.CursorLocation = adUseClient
    rdoEntraNota.Open SQL, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
    
   ' Set rdoEntraNota = conexao.OpenResultset(SQL, Options:=rdExecDirect)
    
    If Not rdoEntraNota.EOF Then
       If IsNull(rdoEntraNota("CV_UltimoDiaMes")) Then
          Data = Format(Date, "dd/mm/yyyy")
       Else
          Data = Format(rdoEntraNota("CV_UltimoDiaMes"), "dd/mm/yyyy")
       End If
    End If
    
    rdoEntraNota.Close
    txtDataEntrada.Text = Data
    
End Sub

Private Sub clsGravaItem_GravacaoErro(ByVal ErroNumero As Long, ByVal Descricao As String)

    MsgBox "Erro gravando item da Nota Fiscal. Tente novamente.", vbCritical, "Erro"
    MostraErro

End Sub

Private Sub clsGravaItem_GravacaoOK(ByVal Resultado As String)

    Dim DescricaoDevol As String
    Dim MotivoDevol As String
    Dim rdoDevolucao As New ADODB.Recordset

    On Error Resume Next

    txtTotalCalculado = Format(Resultado, "###,###,###,###0.00")
    
    txtCodigoOperacao.Enabled = False
    cmbNaturezaOperacao.Enabled = False
    
    If txtReferencia.Text = RefDevolucao Then
        
        Screen.MousePointer = 0
        
        DescricaoDevol = ""
        
        'lo
            DescricaoDevol = InputBox("Entre com a descrição para este Produto :", "Produto a Devolver", left(DescricaoDevol, 38))
            If Trim(DescricaoDevol) = "" Then
                MsgBox "Você deve especificar uma descrição para este produto!", vbExclamation, "Atenção"
            ElseIf Len(Trim(DescricaoDevol)) > 38 Then
                MsgBox "Descrição muito longa. O tamanho máximo é de 38 caracteres.", vbInformation, "Informação"
            End If
        Do While Trim(DescricaoDevol) = "" Or Len(Trim(DescricaoDevol)) > 38
                
        Loop
            MotivoDevol = InputBox("Entre com o motivo da devolução para este Produto :", "Produto a Devolver", left(MotivoDevol, 50))
            If Trim(MotivoDevol) = "" Then
                MsgBox "Você deve especificar um motivo para este produto!", vbExclamation, "Atenção"
            ElseIf Len(Trim(MotivoDevol)) > 50 Then
                MsgBox "Motivo de devolução muito extenso. O tamanho máximo é de 50 caracteres.", vbInformation, "Informação"
            End If
        Do While Trim(MotivoDevol) = "" Or Len(Trim(MotivoDevol)) > 50
        
        Loop
        Screen.MousePointer = 11
        
        
        SQL = "ProdutoDevolver " _
                & Val(txtNotaFiscal.Text) & ", '" _
                & txtSerie.Text & "', " _
                & Val(txtNomeFantasia.Text) & ", '" _
                & DescricaoDevol & "', '" _
                & MotivoDevol & "'"
                
        rdoDevolucao.CursorLocation = adUseClient
        rdoDevolucao.Open SQL, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
        
       ' Set rdoDevolucao = rdoCnSupBatch.OpenResultset("ProdutoDevolver " _
                & Val(txtNotaFiscal.Text) & ", '" _
                & txtSerie.Text & "', " _
                & Val(txtNomeFantasia.Text) & ", '" _
                & DescricaoDevol & "', '" _
                & MotivoDevol & "'", Options:=rdExecDirect)
        
        grdItens.AddItem RefDevolucao & Chr(vbKeyTab) & "0" & Chr(vbKeyTab) _
        & DescricaoDevol & Chr(vbKeyTab) & txtQuantidade.Text _
        & Chr(vbKeyTab) & Format(txtPrecoUnitario.Text, "###,###,###,###0.00") & Chr(vbKeyTab) _
        & Format(txtAliquotaIPI.Text, "0.00") & Chr(vbKeyTab) & Format(txtPercentualDesconto.Text, "0.00") _
        & Chr(vbKeyTab) & Format((CDbl(Format(txtPrecoUnitario.Text, "0.00")) + (1 - (CDbl(Format(txtPercentualDesconto.Text, "0.00")) / 100))) + (1 + (CDbl(Format(txtAliquotaIPI.Text, "0.00")) / 100)), "###,###,###,###0.00") + Val(txtQuantidade.Text) _
        & Chr(vbKeyTab) & "0"
        
        grdItens.RowData(grdItens.Rows - 1) = rdoDevolucao(0)
        
        BateNota = CDbl(Format(txtTotalCalculado.Text, "0.00")) - CDbl(Format(txtValorTotalNota.Text, "0.00"))
        txtBateNota.Text = Format(BateNota, "###,###,###0.00")
        
        clsItensNota.TornaLinhaVisivel grdItens.Rows - 1
        
        Screen.MousePointer = 0

    Else
        SobeDados
    End If
    
    clsGravaItem.Limpar
    
    lblPedido.Visible = True
    lblComprador.Visible = True
    lblCFOP.Visible = True
    lblFrete.Visible = True
    lblEntrega.Visible = True
    lblEntrega2.Visible = True
    lblCondPagto.Visible = True
    lblCondPagto2.Visible = True
    

End Sub

Sub SobeDados()

    grdItens.TextMatrix(grdItens.Row, 0) = txtReferencia.Text
    grdItens.TextMatrix(grdItens.Row, 1) = txtNumeroPedido.Text
    grdItens.TextMatrix(grdItens.Row, 3) = txtQuantidade.Text
    grdItens.TextMatrix(grdItens.Row, 4) = Format(txtPrecoUnitario.Text, "###,###,###,###0.00")
    'grdItens.TextMatrix(grdItens.Row, 5) = Format(txtAliquotaIPI.Text, "0.00")
    'grdItens.TextMatrix(grdItens.Row, 6) = Format(txtPercentualDesconto.Text, "0.00")
    'grdItens.TextMatrix(grdItens.Row, 7) = Format((CDbl(Format(txtPrecoUnitario.Text, "0.00")) + (1 - (CDbl(Format(txtPercentualDesconto.Text, "0.00")) / 100))) + (1 + (CDbl(Format(txtAliquotaIPI.Text, "0.00")) / 100)) + Val(txtQuantidade.Text), "###,###,###,###0.00")
    grdItens.TextMatrix(grdItens.Row, 7) = Format((CDbl(Format(txtPrecoUnitario.Text, "0.00")) * (1 - (CDbl(Format(txtPercentualDesconto.Text, "0.00")) / 100))) * (1 + (CDbl(Format(txtAliquotaIPI.Text, "0.00")) / 100)) * Val(txtQuantidade.Text), "###,###,###,###0.00")
    'grdItens.TextMatrix(grdItens.Row, 7) = Format((txtPrecoUnitario.Text * txtQuantidade.Text), "0.00")
    
    BateNota = CDbl(Format(txtTotalCalculado.Text, "0.00")) - CDbl(Format(txtValorTotalNota.Text, "0.00"))
    txtBateNota.Text = Format(BateNota, "###,###,###0.00")
    
   'rdoCnSupBatch.Execute "Exec CriticaItensEntradaCompras " & txtNotaFiscal.Text & ", '" & txtSerie.Text & "', " & Val(txtNomeFantasia.Text) & ",'" & txtReferencia.Text & "'"
    Call LerItemCritica
        
End Sub

Private Sub clsGravaItem_VerificacaoErro(ByVal Controle As String)

    On Error Resume Next

    MsgBox "Dados inválidos ou insuficientes!", vbExclamation, "Atenção"
    Me.Controls(Controle).SetFocus
    
    Err.Clear

End Sub

Private Sub clsGravaItem_VerificacaoOK()

    Dim TipoAtualizacao As Long

    If NotaJaExiste And Not SoExclusao Then
        If txtReferencia.Text = grdItens.TextMatrix(grdItens.Row, 0) Or txtReferencia.Text = RefDevolucao Then
            If CDbl(Format(txtPrecoUnitario.Text, "0.00")) <= 0 Then
                MsgBox "Preco Unitário inválido!", vbExclamation, "Atenção"
                txtPrecoUnitario.SetFocus
                Exit Sub
            End If
            
            If Val(txtQuantidade.Text) <= 0 Then
                MsgBox "Quantidade inválida!", vbExclamation, "Atenção"
                txtQuantidade.SetFocus
                Exit Sub
            End If
            
            If txtReferencia.Text <> RefDevolucao Then
                If Val(txtQuantidade.Text) > Val(grdItens.TextMatrix(grdItens.Row, 8)) Then
                    MsgBox "A quantidade a entrar não pode ser maior que o saldo do pedido!", vbExclamation, "Atenção"
                    txtQuantidade.SetFocus
                    Exit Sub
                End If
            End If
            
            Screen.MousePointer = 11
            
            If clsGravaItem.Adicionar Or txtReferencia.Text = RefDevolucao Then
                TipoAtualizacao = 1
            Else
                TipoAtualizacao = 0
            End If
            
            clsGravaItem.ExtraValues = Val(txtNotaFiscal.Text) & "; '" & UCase(txtSerie.Text) & "'; " & Val(txtNomeFantasia.Text)
            clsGravaItem.ExecProcedure = True
            clsGravaItem.SQL = "EntradaNFCompra " & TipoAtualizacao & ", "
            
            clsGravaItem.Gravar
            Screen.MousePointer = 0
        Else
            MsgBox "A referência digitada não confere com a selecionada!", vbExclamation, "Atenção"
        End If
    Else
        MsgBox "Você deve primeiro gravar a capa da Nota Fiscal antes de gravar os itens!", vbExclamation, "Atenção"
    End If

End Sub

Private Sub clsItensNota_Acabou()

    grdItens.RowData(grdItens.Rows - 1) = UltimoCFO

    If grdItens.Rows > 2 And grdItens.TextMatrix(1, 0) = "" Then
        grdItens.RemoveItem 1
    End If
    
    grdItens.Redraw = True

End Sub

Private Sub clsItensNota_BeforeAdd(rdoInterno As Variant)

    If LinhaCFO > -1 Then
        grdItens.RowData(grdItens.Rows - 1) = UltimoCFO
    End If
    
    LinhaCFO = LinhaCFO + 1
    
    If rdoInterno("PI_Referencia") <> RefDevolucao Then
        UltimoCFO = (rdoInterno("CodigoOperacao") + 1000) + rdoInterno("NaturezaOperacao")
    Else
        UltimoCFO = rdoInterno("CodigoOperacao")
    End If

End Sub

Private Sub clsNotaFiscal_GravacaoErro(ByVal ErroNumero As Long, ByVal Descricao As String)
    
    MsgBox "Erro gravando a capa da Nota Fiscal. Tente novamente.", vbCritical, "Erro"
    MostraErro

End Sub

Private Sub clsNotaFiscal_GravacaoOK(ByVal Resultado As String)

    Dim SalvaNota As Long
    Dim SalvaForne As String
    Dim SalvaSerie As String
    Dim SalvaCGC As String

    If Not NotaJaExiste Then
        MsgBox "Gravação concluída com sucesso." & Chr(13) & "Obs.: Serão removidos, agora, os ítens que possuem natureza de operação diferente da selecionada.", vbInformation, "Informação"
        FiltraCodigoOperacao
        txtCodigoOperacao.Tag = txtCodigoOperacao.Text
        cmbNaturezaOperacao.Tag = Val(cmbNaturezaOperacao.Text)
    End If
    
    If txtCodigoOperacao.Text = txtCodigoOperacao.Tag And Val(cmbNaturezaOperacao.Text) = Val(cmbNaturezaOperacao.Tag) Then
'        If GrdItens.Rows = 2 And GrdItens.TextMatrix(1, 0) = "" Then
'            txtCodigoOperacao.Enabled = True
'            cmbNaturezaOperacao.Enabled = True
'        Else
'            txtCodigoOperacao.Enabled = False
'            cmbNaturezaOperacao.Enabled = False
'        End If
        
        NotaJaExiste = True
        
        cmdVencimentos.Enabled = True
        
        grdItens.Enabled = True
        pnlItens.Enabled = True
        
        grdItens.Row = 1
        grdItens.Col = 0
        grdItens.ColSel = grdItens.Cols - 1
        grdItens.SetFocus
        DesceDados
        PreencheCapa
        txtQuantidade.SetFocus
    Else
        MsgBox "Houve uma alteração de Código de Operação e/ou Natureza de Operação, a Nota Fiscal será reexibida.", vbInformation, "Informação"
        
        Screen.MousePointer = 11
    
        SalvaNota = Val(txtNotaFiscal.Text)
        SalvaForne = txtNomeFantasia.Text
        SalvaSerie = txtSerie.Text
        SalvaCGC = txtCGC.Text
        
        ResetMe
        
        txtCGC.Text = SalvaCGC
        txtNomeFantasia.Text = SalvaForne
        txtNotaFiscal.Text = SalvaNota
        txtSerie = SalvaSerie
        
        CarregaNotaFiscal
        
        Screen.MousePointer = 0
    End If

End Sub

Private Function LocalizaTipoNatureza(ByVal Codigo As Long) As String

    Dim MaximoNatureza As Long
    Dim Indice As Long
    Dim TipoNatureza As String
    
    MaximoNatureza = UBound(matNatureza)
    
    For Indice = 0 To MaximoNatureza Step 1
        If matNatureza(Indice).CFO = Codigo Then
            Exit For
        End If
    Next Indice
    
    If Indice <= MaximoNatureza Then
        LocalizaTipoNatureza = matNatureza(Indice).TipoNatureza
    Else
        LocalizaTipoNatureza = ""
    End If

End Function

Private Sub FiltraCodigoOperacao()

    Dim Linha As Long
    Dim CFO As Long
    Dim CFOAux As Long
    Dim Maximo As Long
    Dim TipoNatureza As String
    
    On Error Resume Next
    
    Err.Clear
    
    Maximo = UBound(matCFO)
    
    CFO = Val(txtCodigoOperacao.Text)
    CFOAux = 0
    
    For Linha = 0 To Maximo Step 1
        If CFO = matCFO(Linha).CFO Then
            CFOAux = matCFO(Linha).CFOAux
            
            Exit For
        End If
    Next Linha
    
    CFO = (CFOAux * 1000) + Val(cmbNaturezaOperacao.Text)
    
    TipoNatureza = LocalizaTipoNatureza(CFO)
    
    Maximo = grdItens.Rows - 1
    
    For Linha = Maximo To 1 Step -1
        If grdItens.RowData(Linha) <> 0 And grdItens.TextMatrix(Linha, 14) <> CFOAux And grdItens.TextMatrix(Linha, 0) <> RefDevolucao Then
            If TipoNatureza <> "" Then
                If TipoNatureza <> LocalizaTipoNatureza(CFOAux) Then
                    grdItens.RemoveItem Linha
                End If
            Else
                grdItens.RemoveItem Linha
            End If
        End If
    Next Linha
    
    If Err.Number <> 0 Then
        Err.Clear
        grdItens.AddItem ""
        grdItens.RemoveItem 1
    End If

End Sub

Private Sub clsNotaFiscal_LeituraOK()
    
    Dim SalvaNomeFantasia As String
    Dim I As Integer
    
    SalvaNomeFantasia = txtNomeFantasia.Text

    clsNotaFiscal.Preencher
    
    clsItensNota.Clear
    'clsItensNota.SQL = "Select CI_Observacao, 0 as CodigoOperacao, PI_NumeroPedido, PI_Referencia, " _
                       & "PR_Descricao, CI_Quantidade, CI_PrecoUnitario, " _
                       & "CI_AliquotaIPI, CI_PercentualDesconto, PI_SaldoPedido, " _
                       & "(CI_PrecoUnitario * (1 - (CI_PercentualDesconto/100.0))) * (1 + (CI_AliquotaIPI/100.0)) * CI_Quantidade as Total, 0 as NaturezaOperacao " _
                       & "From Produto, ItemPedido, ItemNFCompra " _
                       & "where PR_Referencia = CI_Referencia and " _
                       & "CI_NossoPedido = PI_NumeroPedido and CI_Referencia = PI_Referencia and " _
                       & "PI_Filial = '" & cmbLocal.Text & "' and " _
                       & "CI_NotaFiscal = " & Val(txtNotaFiscal.Text) & " and " _
                       & "CI_Serie = '" & Trim(txtSerie.Text) & "' and " _
                       & "CI_Fornecedor = " & Val(txtNomeFantasia.Text) _
                       & " Union( Select CI_Observacao,CI_Item as CodigoOperacao, 0, '" & RefDevolucao & "', " _
                       & "DE_Descricao as PR_Descricao, CI_Quantidade, CI_PrecoUnitario, " _
                       & "CI_AliquotaIPI, CI_PercentualDesconto, 0, " _
                       & "(CI_PrecoUnitario * (1 - (CI_PercentualDesconto/100.0))) * (1 + (CI_AliquotaIPI/100.0)) * CI_Quantidade as Total, 0 " _
                       & "From ItemNFCompra, DescricaoEspecial " _
                       & "where CI_NossoPedido = 0 and CI_NotaFiscal = DE_NotaFiscalEntrada and " _
                       & "CI_Serie = DE_SerieEntrada and CI_Fornecedor = DE_FornecedorEntrada and " _
                       & "CI_Item = DE_Item and " _
                       & "CI_NotaFiscal = " & Val(txtNotaFiscal.Text) & " and " _
                       & "CI_Serie = '" & Trim(txtSerie.Text) & "' and " _
                       & "CI_Fornecedor = " & Val(txtNomeFantasia.Text) & ") " _
                       & "order by PI_NumeroPedido, PR_Descricao "
    clsItensNota.SQL = "Select PC_LojaReserva,PR_Residencia,CO_Nome as Comprador,PC_DespesaFrete as Frete,PC_SemanaEntrega as Entrega,CP_Descricao as Condpag,CI_observacao, PC_CodigoOperacao as CodigoOperacao, PI_NumeroPedido, PI_Referencia, " _
                       & "PR_Descricao, CI_Quantidade, CI_PrecoUnitario, " _
                       & "CI_AliquotaIPI, CI_PercentualDesconto, PI_SaldoPedido, " _
                       & "(CI_PrecoUnitario * (1 - (CI_PercentualDesconto/100.0))) * (1 + (CI_AliquotaIPI/100.0)) * CI_Quantidade as Total, 0 as NaturezaOperacao " _
                       & "From CondicaoPagto,Produto,Comprador,CapaPedido,ItemNFCompra, DescricaoEspecial ,ItemPedido " _
                       & "where CP_VendaCompra = 'C' and CP_CodigoCondicao = PC_CondicaoPagamento and PC_NumeroPedido = PI_NumeroPedido and PC_CodigoComprador = CO_CodigoComprador and PR_Referencia = CI_Referencia and " _
                       & "CI_NossoPedido = PI_NumeroPedido and CI_Referencia = PI_Referencia and " _
                       & "PI_Filial = '" & cmbLocal.Text & "' and " _
                       & "CI_NotaFiscal = " & Val(txtNotaFiscal.Text) & " and " _
                       & "CI_Serie = '" & Trim(txtSerie.Text) & "' and " _
                       & "CI_Fornecedor = " & Val(txtNomeFantasia.Text) _
                       & " Union( Select PC_LojaReserva,PR_Residencia,CO_Nome as Comprador,PC_DespesaFrete as Frete,PC_SemanaEntrega as Entrega,CP_Descricao as Condpag,CI_Observacao,CI_Item as CodigoOperacao, 0, '" & RefDevolucao & "', " _
                       & "DE_Descricao as PR_Descricao, CI_Quantidade, CI_PrecoUnitario, " _
                       & "CI_AliquotaIPI, CI_PercentualDesconto, 0, " _
                       & "(CI_PrecoUnitario * (1 - (CI_PercentualDesconto/100.0))) * (1 + (CI_AliquotaIPI/100.0)) * CI_Quantidade as Total, 0 " _
                       & "From Produto,CondicaoPagto,Comprador,CapaPedido,ItemNFCompra, DescricaoEspecial ,ItemPedido " _
                       & "where CP_VendaCompra = 'C' and CP_CodigoCondicao = PC_CondicaoPagamento and CI_NossoPedido = 0 and CI_NotaFiscal = DE_NotaFiscalEntrada and " _
                       & "CI_Serie = DE_SerieEntrada and CI_Fornecedor = DE_FornecedorEntrada and " _
                       & "CI_Item = DE_Item and " _
                       & "CI_NotaFiscal = " & Val(txtNotaFiscal.Text) & " and " _
                       & "CI_Serie = '" & Trim(txtSerie.Text) & "' and " _
                       & "CI_Fornecedor = " & Val(txtNomeFantasia.Text) & ") " _
                       & "order by PI_NumeroPedido, PR_Descricao "

    LinhaCFO = -1
    clsItensNota.Preencher
    
    For I = 1 To grdItens.Rows - 1
        If grdItens.TextMatrix(I, 9) <> "" Then
            grdItens.Row = I
            grdItens.RowSel = I
            grdItens.ColSel = 9
            grdItens.FillStyle = flexFillRepeat
            grdItens.CellForeColor = &HDD&
            grdItens.FillStyle = flexFillSingle
        Else
            grdItens.Row = I
            grdItens.RowSel = I
            grdItens.ColSel = 9
            grdItens.FillStyle = flexFillRepeat
            grdItens.CellForeColor = &H0&
            grdItens.FillStyle = flexFillSingle
        End If
    Next I

    
    txtNomeFantasia.Text = SalvaNomeFantasia
    
End Sub

Sub RemoveRepetidos()

    Dim Linha As Long
    Dim Maximo As Long
    Dim QtdAnterior As String
    Dim Referencia As String
    Dim Pedido As String
    Dim I As Integer
    
    Maximo = grdItens.Rows - 1

    If grdItens.Rows > 2 Then
        grdItens.Redraw = False
        
        grdItens.Row = 1
        grdItens.Col = 1
        grdItens.ColSel = 0
        grdItens.Sort = flexSortNumericAscending
        
        QtdAnterior = "0"
        Referencia = ""
        Pedido = ""
        For Linha = Maximo To 1 Step -1
            If grdItens.TextMatrix(Linha, 1) <> Pedido Or grdItens.TextMatrix(Linha, 0) <> Referencia Then
                Pedido = grdItens.TextMatrix(Linha, 1)
                Referencia = grdItens.TextMatrix(Linha, 0)
                QtdAnterior = Trim(grdItens.TextMatrix(Linha, 3))
            ElseIf grdItens.TextMatrix(Linha, 1) = Pedido And grdItens.TextMatrix(Linha, 0) = Referencia Then
                If grdItens.TextMatrix(Linha, 0) <> RefDevolucao Then
                    If QtdAnterior = "" Then
                        grdItens.RemoveItem Linha + 1
                    Else
                        grdItens.RemoveItem Linha
                    End If
                End If
            End If
        Next Linha
        
        grdItens.Row = 1
        grdItens.Col = 1
        grdItens.ColSel = 2
        grdItens.Sort = flexSortGenericAscending
        
        grdItens.Row = 1
        grdItens.Col = 0
        grdItens.ColSel = grdItens.Cols - 1
        grdItens.Sort = flexSortNone
        
        grdItens.Redraw = True
    End If
    
    For I = 1 To grdItens.Rows - 1
        If grdItens.TextMatrix(I, 9) = grdItens.TextMatrix(I, 8) Then
            grdItens.TextMatrix(I, 9) = ""
        End If
    Next I

End Sub

Private Sub clsNotaFiscal_PreenchimentoOK()
    
    MontaComboNatureza Val(txtCodigoOperacao.Text), cmbNaturezaOperacao
        
    cmbNaturezaOperacao.ListIndex = GetIndiceVal(lblnaturezaoperacao.Caption, cmbNaturezaOperacao)

    Pesquisou = True
    
    txtCodigoOperacao.Enabled = False
    cmbNaturezaOperacao.Enabled = False
    
    txtCodigoOperacao.Tag = txtCodigoOperacao.Text
    cmbNaturezaOperacao.Tag = Val(cmbNaturezaOperacao.Text)
    
    cmbNaturezaOperacao.Enabled = False

End Sub

Private Sub clsNotaFiscal_RegistroEncontrado(rdoInterno As Variant, Cancelar As Boolean)

    SoExclusao = False

    NotaJaExiste = True
    
    If rdoInterno("CC_Situacao") = "E" Then
        grdItens.Enabled = False
        pnlItens.Enabled = False
        SoExclusao = True
        MsgBox "Esta nota não pode ser editada aqui pois já foi Encerrada!", vbInformation, "Informação"
    ElseIf rdoInterno("CC_Situacao") <> "D" And rdoInterno("CC_Situacao") <> "E" Then
        SoExclusao = False
        Cancelar = True
        MsgBox "Esta nota não pode ser visualizada aqui pois já foi Liberada!", vbInformation, "Informação"
        ResetMe
        txtCGC.SetFocus
        Exit Sub
    Else
        grdItens.Enabled = True
        pnlItens.Enabled = True
    End If
    
    If SoExclusao Then
        cmdEntraNota(1).Enabled = False
        cmdEntraNota(3).Enabled = False
    End If
    
    txtCGC.Enabled = False
    txtNotaFiscal.Enabled = False
    txtSerie.Enabled = False
    cmbLocal.Enabled = False
    
    TotalCalculado = rdoInterno("CC_ValorCalculado")
    TotalNota = rdoInterno("CC_ValorTotalNota")
    Frete = rdoInterno("CC_Frete")
    Embalagem = rdoInterno("CC_Embalagem")
    Despesas = rdoInterno("CC_Despesas")
    Juros = rdoInterno("CC_Juros")
    Outros = rdoInterno("CC_Outros")
    
    BateNota = TotalCalculado - TotalNota
    
    txtBateNota.Text = Format(BateNota, "###,###,###0.00")
    
End Sub

Private Sub clsNotaFiscal_RegistroNaoExiste()

    Dim SalvaCGC As String
    Dim SalvaNomeFantasia As String
    Dim SalvaNotaFiscal As String
    Dim SalvaSerie As String
    Dim SalvaLocal As String
    
    SalvaCGC = txtCGC.Text
    SalvaNomeFantasia = txtNomeFantasia.Text
    SalvaNotaFiscal = txtNotaFiscal.Text
    SalvaSerie = UCase(txtSerie.Text)
    SalvaLocal = cmbLocal.Text
    
    ResetMe
    
    txtCGC.Text = SalvaCGC
    txtNomeFantasia.Text = SalvaNomeFantasia
    txtNotaFiscal.Text = SalvaNotaFiscal
    txtSerie.Text = SalvaSerie
    cmbLocal.ListIndex = GetIndice(SalvaLocal, cmbLocal)

    Pesquisou = True
    NotaJaExiste = False
    SoExclusao = False
    
    cmdEntraNota(1).Enabled = True
    cmdEntraNota(3).Enabled = True
    
    grdItens.Enabled = False
    pnlItens.Enabled = False

End Sub

Private Sub clsNotaFiscal_VerificacaoErro(ByVal Controle As String)
    
    On Error Resume Next
    
    MsgBox "Dados inválidos ou insuficientes!", vbExclamation, "Atenção"
    
    If UCase(Controle) <> UCase("lblNaturezaOperacao") Then
        Me.Controls(Controle).SetFocus
    Else
        If txtCodigoOperacao <> "" Then
            cmbNaturezaOperacao.SetFocus
        Else
            txtCodigoOperacao.SetFocus
        End If
    End If
    
    Err.Clear

End Sub

Private Sub clsNotaFiscal_VerificacaoOK()

    If Val(ConverteVirgula(Format(txtValorTotalNota.Text, "0.00"))) = 0 Then
        Screen.MousePointer = 0
        
        MsgBox "Informe o total da Nota!", vbExclamation, "Atenção"
        
        Exit Sub
    End If

    If VerificaDatas Then
        Screen.MousePointer = 11
        
'        If AcertoConsignacao Then
'            Screen.MousePointer = 0
'            frmNotaOrigem.Hide
'
'            frmNotaOrigem.txtNotaFiscal.Text = IIf(Val(lblNotaOrigem.Caption) = 0, "", Val(lblNotaOrigem.Caption))
'            frmNotaOrigem.txtSerie.Text = lblSerieOrigem.Caption
'
'            frmNotaOrigem.Show 1
'
'            If frmNotaOrigem.Cancelou Then
'                Unload frmNotaOrigem
'
'                Exit Sub
'            End If
'
'            'lblNotaOrigem.Caption = frmNotaOrigem.txtNotaFiscal.Text
'            'lblSerieOrigem.Caption = frmNotaOrigem.txtSerie.Text
'
'            clsNotaFiscal.Re_Get
'
'            Unload frmNotaOrigem
'        Else
'            lblNotaOrigem.Caption = "0"
'            lblSerieOrigem.Caption = ""
'        End If


    
        Screen.MousePointer = 11
        
        clsNotaFiscal.Adicionar = Not NotaJaExiste
        clsNotaFiscal.UseWhere = True
        clsNotaFiscal.SQL = "CapaNFCompra"
        clsNotaFiscal.ClausulaWhere = "where CC_NotaFiscal = " & txtNotaFiscal.Text & " and CC_Serie = '" & txtSerie.Text & "' and CC_Fornecedor = " & Val(txtNomeFantasia.Text)

        lblNotaOrigem.Caption = "0"
        lblSerieOrigem.Caption = ""

        clsNotaFiscal.Gravar
        
        Screen.MousePointer = 0
    End If

End Sub

'Private Function AcertoConsignacao() As Boolean
'
'    Dim CFO As Long
'    Dim Indice As Long
'    Dim Maximo As Long
'
'    AcertoConsignacao = False
'
'    If Val(cmbNaturezaOperacao.Text) = 3 Or Val(cmbNaturezaOperacao.Text) = 4 Then
'        Maximo = UBound(matNatureza)
'        CFO = Val(txtCodigoOperacao.Text)
'
'        For Indice = 0 To Maximo Step 1
'            If CFO = matNatureza(Indice).CFO And matNatureza(Indice).TipoNatureza = "V" Then
'                AcertoConsignacao = True
'
'                Exit Function
'            End If
'        Next Indice
'    End If
'
'End Function

Private Function VerificaDatas() As Boolean

    VerificaDatas = False
    
    If DateDiff("d", mskDataEmissao.Text, txtDataEntrada.Text) >= 0 Then
        If DateDiff("d", mskDataEmissao.Text, mskDataRecebimento.Text) >= 0 Then
            If DateDiff("d", mskDataRecebimento.Text, txtDataEntrada.Text) >= 0 Then
                VerificaDatas = True
            Else
                MsgBox "A data de Entrada deve ser maior ou igual à data de Recebimento!", vbExclamation, "Atenção"
            End If
        Else
            MsgBox "A data de Recebimento deve ser maior ou igual à data de Emissao!", vbExclamation, "Atenção"
        End If
    Else
        MsgBox "A data de Entrada deve ser maior ou igual à data de Emissão!", vbExclamation, "Atenção"
    End If

End Function

Private Sub cmbNaturezaOperacao_Click()
    
    lblnaturezaoperacao.Caption = cmbNaturezaOperacao.Text

End Sub

Function VerNaturezaOperacao() As Boolean

    Dim Indice As Long
    Dim Maximo As Long

    VerNaturezaOperacao = True
    
    Maximo = UBound(matNatureza)
    
    If Val(cmbNaturezaOperacao.Text) = 3 Or Val(cmbNaturezaOperacao.Text) = 4 Then
        For Indice = 0 To Maximo Step 1
            If Val(txtCodigoOperacao.Text) = matNatureza(Indice).CFO And matNatureza(Indice).TipoNatureza = "V" Then
                Exit For
            End If
        Next Indice
        
        If Indice <= Maximo Then
        
        End If
    End If

End Function

Private Sub cmdEntraNota_Click(Index As Integer)
    
    Select Case Index
        Case 0
        
            lblPedido.Visible = False
            lblComprador.Visible = False
            lblCFOP.Visible = False
            lblFrete.Visible = False
            lblEntrega.Visible = False
            lblEntrega2.Visible = False
            lblCondPagto.Visible = False
            lblCondPagto2.Visible = False
        
            Unload Me
        Case 1
            If Not txtCGC.Enabled Then
                If VerNaturezaOperacao Then
                   WCodigoOperacaoNovo = txtCodigoOperacao.Text
                   txtCodigoOperacao.Text = WCodigoOperacaoVelho
                   
                   clsNotaFiscal.Verificar
                    
                   txtCodigoOperacao.Text = WCodigoOperacaoNovo
                End If
            End If
        Case 2
            If NotaJaExiste Then
                If MsgBox("Confirma a exclusão desta nota?", vbYesNo + vbQuestion + vbDefaultButton2, "Exclusão de Nota") = vbYes Then
                    ExcluiNota
                End If
            End If
        Case 3
            If NotaJaExiste Then
                If ExisteVencimentos Or (Trim(cmbLocal.Text) = "CMCE" Or Trim(cmbLocal.Text) = "MC85E" And LocalizaTipoNatureza(Val(txtCodigoOperacao.Text)) = "RCt") Then
                    If MsgBox("Confirma o encerramento desta nota?", vbYesNo + vbQuestion + vbDefaultButton2, "Encerramento") = vbYes Then
                        EncerraNota
                    End If
                Else
                    MsgBox "Ainda não foram especificados os Vencimentos para esta Nota Fiscal. A mesma não pode ser encerrada.", vbExclamation, "Atenção"
                End If
            End If
        Case 4
            ResetMe
    End Select

End Sub

Private Function ExisteVencimentos() As Boolean

    Dim rdoVencimentos As New ADODB.Recordset
    
    On Error Resume Next
    
    Set rdoVencimentos = rdoCnSup.OpenResultset("Select VF_Parcela From " _
        & "VencimentosFornecedor Where VF_NotaFiscal = " & Val(txtNotaFiscal.Text) _
        & " and VF_Serie = '" & txtSerie.Text & "' and VF_Fornecedor = " _
        & Val(txtNomeFantasia.Text), Options:=rdExecDirect)
        
    If Not rdoVencimentos.EOF Then
        ExisteVencimentos = True
    Else
        ExisteVencimentos = False
    End If

    rdoVencimentos.Close

End Function

Private Sub EncerraNota()

    Dim rdoEntraNota As New ADODB.Recordset

    

    Set rdoEntraNota = rdoCnSupBatch.OpenResultset("Select PA_LimiteBateNota from Parametros", Options:=rdExecDirect)
    If CDbl(Format(txtBateNota.Text, "0.00")) > rdoEntraNota("PA_LimiteBateNota") Then
        rdoEntraNota.Close
        
        If MsgBox("Bate-nota é maior que o limite. Confirma o encerramento da nota?", vbYesNo + vbDefaultButton2 + vbQuestion, "Encerramento de Nota") = vbYes Then
            Encerramento
        End If
    Else
        rdoEntraNota.Close
        Encerramento
    End If
    
End Sub

Sub Encerramento()

    Dim rdoItens As New ADODB.Recordset
    Dim CriticaInt As Integer
    Dim VarF As Integer
    
    I = 0
    
    On Error Resume Next

    Screen.MousePointer = 11

    Set rdoItens = rdoCnSupBatch.OpenResultset("Select Count(*) From ItemNFCompra where CI_NotaFiscal = " & txtNotaFiscal.Text & " and CI_Serie = '" & txtSerie.Text & "' and CI_Fornecedor = " & Val(txtNomeFantasia.Text), Options:=rdExecDirect)
    
    If rdoItens(0) = 0 Then
        rdoItens.Close
        MsgBox "Ainda não existem ítens para esta nota, por isso não pode ser encerrada.", vbExclamation, "Atenção"
        
        Screen.MousePointer = 0
        
        Exit Sub
    Else
        rdoItens.Close
        
        SQL = ""
'        SQL = "UPDATE CapaNFCompra Set cc_TipoEntrada = pc_tipopedido FROM CapaPedido, ItemnfCompra, CapaNFCompra " & _
'              "Where ci_NossoPedido = pc_NumeroPedido and cc_NotaFiscal = ci_NotaFiscal and cc_Serie = ci_Serie and " & _
'              "cc_Fornecedor = ci_Fornecedor and cc_NotaFiscal = " & txtNotaFiscal.Text & " and cc_Serie = '" & txtSerie.Text & "' and " & _
'              "cc_Fornecedor = " & Val(txtNomeFantasia.Text) & " and pc_NumeroPedido = " & grdItens.TextMatrix(grdItens.Row, 1) & " and " & _
'              "cc_DataEntrada = '" & Format(txtDataEntrada.Text, "yyyy/mm/dd") & "'"

        SQL = "UPDATE CapaNFCompra Set cc_TipoEntrada = pc_tipopedido, cc_AcaoEntrada = pc_AcaoPedido FROM CapaPedido, ItemnfCompra, CapaNFCompra " & _
              "Where ci_NossoPedido = pc_NumeroPedido and cc_NotaFiscal = ci_NotaFiscal and cc_Serie = ci_Serie and " & _
              "cc_Fornecedor = ci_Fornecedor and cc_NotaFiscal = " & txtNotaFiscal.Text & " and cc_Serie = '" & txtSerie.Text & "' and " & _
              "cc_Fornecedor = " & Val(txtNomeFantasia.Text) & " and cc_DataEntrada = '" & Format(txtDataEntrada.Text, "yyyy/mm/dd") & "'"

          rdoCnSupBatch.Execute (SQL)
        
        If CriticaFrete Then
            If Trim(cmbLocal.Text) = "CMCE" And LocalizaTipoNatureza(Val(txtCodigoOperacao.Text)) = "RCt" Then
                rdoCnSupBatch.Execute "Exec EncerraNotaFiscalCompra " & txtNotaFiscal.Text & ", '" & txtSerie.Text & "', " & Val(txtNomeFantasia.Text) & ", 1", rdExecDirect
            ElseIf Trim(cmbLocal.Text) = "MC85E" And LocalizaTipoNatureza(Val(txtCodigoOperacao.Text)) = "RCt" Then
                   rdoCnSupBatch.Execute "Exec EncerraNotaFiscalCompra " & txtNotaFiscal.Text & ", '" & txtSerie.Text & "', " & Val(txtNomeFantasia.Text) & ", 2", rdExecDirect
            Else
                   rdoCnSupBatch.Execute "Exec EncerraNotaFiscalCompra " & txtNotaFiscal.Text & ", '" & txtSerie.Text & "', " & Val(txtNomeFantasia.Text) & ", 0", rdExecDirect
            End If
        Else
            Screen.MousePointer = 0
            
            txtFrete.SetFocus
            txtFrete.SelStart = 0
            txtFrete.SelLength = Len(txtFrete.Text)
            
            Exit Sub
        End If
    End If
    
    If Err.Number = 0 Then
        For I = 1 To grdItens.Rows - 1
            If grdItens.TextMatrix(I, 9) <> "" Then
                CriticaInt = CriticaInt + 1
            End If
        Next I
        
        Fornecedor = ""
        For VarF = 1 To Len(txtNomeFantasia.Text)
            If Mid(txtNomeFantasia.Text, VarF, 1) = " " Then
                Exit For
            Else
                Fornecedor = Fornecedor & Mid(txtNomeFantasia.Text, VarF, 1)
            End If
        Next VarF
    
        If Trim(cmbLocal.Text) <> MercConcerto And LocalizaTipoNatureza(Val(txtCodigoOperacao.Text)) <> "RCt" Then
            If (txtValorTotalNota.Text - txtTotalCalculado.Text) < 1.01 And CriticaInt = 0 Then
                lblControlaTela.Caption = "frmEntradaNotaFiscal"
                OutraTela = "S"
                frmLiberaMerc.txtNotaFiscal = txtNotaFiscal.Text
                frmLiberaMerc.txtSerie = txtSerie.Text
                frmLiberaMerc.Txtfornecedor = Numeros(txtNomeFantasia.Text)
                frmLiberaMerc.txtSerie.SetFocus
                frmLiberaMerc.Show
                For I = 1 To grdItens.Rows - 1
                    If Trim(grdItens.TextMatrix(I, 3)) <> "" Then
                        If grdItens.TextMatrix(I, 16) <> "" Then
                            SQL = "SELECT Convert(VarChar(3),CI_Referencia) as Fornecedor FROM ItemNfCompra WHERE " & _
                              "CI_NotaFiscal = " & txtNotaFiscal.Text & " and CI_Serie = '" & txtSerie.Text & "' and " & _
                              "CI_DataEntrada = '" & Format(txtDataEntrada.Text, "mm/dd/yyyy") & "' and CI_Fornecedor = " & Fornecedor & " Group By Convert(VarChar(3),CI_Referencia) " & _
                              "Order By Fornecedor"
                            Set adoChkFornecedor = rdoCnSupBatch.OpenResultset(SQL)
                        '-------------------------------------------------------------------------
                        '--                                                                     --
                        '--   Está rotina ficará inibida até que seja feita a Tela Nova         --
                        '--   que efetuará a distribuição por Nota - Adilson                    --
                        '--                                                                     --
                        '--   Call DistribuicaoAutomatica                                       --
                        '--                                                                     --
                        '-------------------------------------------------------------------------
                        End If
                    End If
                Next I
            ElseIf CriticaInt > 0 Then
                MsgBox "Atenção Nota com Criticas!", vbCritical
                lblControlaTela.Caption = "frmEntradaNotaFiscal"
                frmLiberaMerc.txtNotaFiscal = txtNotaFiscal.Text
                frmLiberaMerc.txtSerie = txtSerie.Text
                frmLiberaMerc.Txtfornecedor = Numeros(txtNomeFantasia.Text)
                frmLiberaMerc.txtSerie.SetFocus
                frmLiberaMerc.Show
                For I = 1 To grdItens.Rows - 1
                    If Trim(grdItens.TextMatrix(I, 3)) <> "" Then
                        If grdItens.TextMatrix(I, 16) <> "" Then
                            SQL = "SELECT Convert(VarChar(3),CI_Referencia) as Fornecedor FROM ItemNfCompra WHERE " & _
                              "CI_NotaFiscal = " & txtNotaFiscal.Text & " and CI_Serie = '" & txtSerie.Text & "' and " & _
                              "CI_DataEntrada = '" & Format(txtDataEntrada.Text, "mm/dd/yyyy") & "' and CI_Fornecedor = " & Fornecedor & " Group By Convert(VarChar(3),CI_Referencia) " & _
                              "Order By Fornecedor"
                            Set adoChkFornecedor = rdoCnSupBatch.OpenResultset(SQL)
                        '-------------------------------------------------------------------------
                        '--                                                                     --
                        '--   Está rotina ficará inibida até que seja feita a Tela Nova         --
                        '--   que efetuará a distribuição por Nota - Adilson                    --
                        '--                                                                     --
                        '--   Call DistribuicaoAutomatica                                       --
                        '--                                                                     --
                        '-------------------------------------------------------------------------
                        End If
                    End If
                Next I
            End If
        End If
        ResetMe
    Else
        MsgBox "Erro encerrando Nota Fiscal. Tente novamente.", vbCritical, "Erro"
        MostraErro
    End If
    
    Screen.MousePointer = 0
    
    CriticaInt = 0
    
End Sub

Private Function CriticaFrete() As Boolean

    Dim rdoFrete As New ADODB.Recordset

    CriticaFrete = True
    
    Set rdoFrete = rdoCnSupBatch.OpenResultset("Select " _
                 & "Sum(PC_DespesaFrete) as Frete " _
                 & "From CapaPedido Where PC_NumeroPedido in (Select " _
                 & "Distinct CI_NossoPedido From ItemNFCompra Where " _
                 & "CI_NotaFiscal = " & Val(txtNotaFiscal.Text) _
                 & " and CI_Serie = '" & txtSerie.Text & "' and " _
                 & "CI_Fornecedor = " & Val(txtNomeFantasia.Text) & ")" _
                 , Options:=rdExecDirect)
    
    If (rdoFrete("Frete") = 0 And CDbl(Format(txtFrete.Text, "0.00")) <> 0) Or (rdoFrete("Frete") <> 0 And CDbl(Format(txtFrete.Text, "0.00")) = 0) Then
        MsgBox "Frete em desacordo com o Pedido. Verificar e entrar em contato com o Compras nota não encerrada.", vbCritical, "Atenção"
        CriticaFrete = False
    End If
    
    rdoFrete.Close

End Function

Private Sub ExcluiNota()

    Dim PodeExcluir As Boolean

    On Error Resume Next

    If SoExclusao Then
        If SituacaoOK Then
            PodeExcluir = True
        Else
            PodeExcluir = False
        End If
    Else
        PodeExcluir = True
    End If

    If PodeExcluir Then
        Screen.MousePointer = 11
    
        rdoCnSupBatch.Execute "Exec EstornaEncerramento " _
             & Val(txtNotaFiscal.Text) & ", " _
             & "'" & txtSerie.Text & "', " _
             & Val(txtNomeFantasia.Text), rdExecDirect
        
        rdoCnSupBatch.BeginTrans
    
        rdoCnSupBatch.Execute "Delete VencimentosFornecedor where VF_NotaFiscal = " & txtNotaFiscal.Text & " and VF_Serie = '" & txtSerie.Text & "' and VF_Fornecedor = " & Val(txtNomeFantasia.Text), rdExecDirect
        rdoCnSupBatch.Execute "Delete ItemNFCompra where CI_NotaFiscal = " & txtNotaFiscal.Text & " and CI_Serie = '" & txtSerie.Text & "' and CI_Fornecedor = " & Val(txtNomeFantasia.Text), rdExecDirect
        rdoCnSupBatch.Execute "Delete CapaNFCompra where CC_NotaFiscal = " & txtNotaFiscal.Text & " and CC_Serie = '" & txtSerie.Text & "' and CC_Fornecedor = " & Val(txtNomeFantasia.Text), rdExecDirect
        rdoCnSupBatch.Execute "Delete DistribuicaoAutomatica WHERE DA_NotaFiscal = " & txtNotaFiscal.Text & " and DA_Serie = '" & txtSerie.Text & "' and DA_CodigoFornecedor = " & Mid(txtNomeFantasia.Text, 1, 3), rdExecDirect
        
        
        If Err.Number = 0 Then
            rdoCnSupBatch.CommitTrans
            ResetMe
        Else
            rdoCnSupBatch.RollbackTrans
            
            MsgBox "Erro excluindo Nota Fiscal. Tente novamente.", vbCritical, "Erro"
            MostraErro
        End If
        
        Screen.MousePointer = 0
    Else
        MsgBox "Esta nota não pode ser excluída, pois possui itens já Liberados.", vbInformation, "Informação"
    End If
    
End Sub

Function SituacaoOK() As Boolean

    Dim rdoSituacao As New ADODB.Recordset
    
    Set rdoSituacao = rdoCnSupBatch.OpenResultset("Select " _
            & "Count(*) as Liberados from ItemNFCompra " _
            & "where CI_NotaFiscal = " & Val(txtNotaFiscal.Text) _
            & " and CI_Serie = '" & txtSerie.Text _
            & "' and CI_Fornecedor = " & Val(txtNomeFantasia.Text) _
            & " and CI_Situacao not in ('D', 'E')", Options:=rdExecDirect)
            
    If rdoSituacao("Liberados") > 0 Then
        SituacaoOK = False
    Else
        SituacaoOK = True
    End If
    
    rdoSituacao.Close

End Function

Sub ResetMe()
    
    lblPedido.Visible = False
    lblComprador.Visible = False
    lblCFOP.Visible = False
    lblFrete.Visible = False
    lblEntrega.Visible = False
    lblEntrega2.Visible = False
    lblEntrega2.Caption = ""
    lblCondPagto.Visible = False
    lblCondPagto2.Visible = False
    lblCondPagto2.Caption = ""
    
    clsNotaFiscal.Limpar
    clsGravaItem.Limpar
    clsItensNota.Clear
    
    cmdEntraNota(1).Enabled = True
    cmdEntraNota(3).Enabled = True
    cmdVencimentos.Enabled = False

    txtCGC.Enabled = True
    txtNotaFiscal.Enabled = True
    txtSerie.Enabled = True
    pnlItens.Enabled = False
    grdItens.Enabled = False
    cmbLocal.Enabled = True
    
    cmbNaturezaOperacao.Clear
    
    txtBateNota.Text = "0,00"
    txtCGC.Text = ""
    txtDataEntrada = Data
    txtPesquisaPedido.Text = ""
    txtPesquisaDescricao.Text = ""
    
    cmbLocal.ListIndex = GetIndice(Almoxarifado, cmbLocal)
    
    Pesquisou = False
    NotaJaExiste = False
    SoExclusao = False
    
    TotalCalculado = 0
    TotalNota = 0
    Frete = 0
    Embalagem = 0
    Despesas = 0
    Juros = 0
    Outros = 0
    BateNota = 0
    
    txtCodigoOperacao.Enabled = True
    cmbNaturezaOperacao.Enabled = True
    
    Screen.MousePointer = 0


End Sub

Private Sub cmdVencimentos_Click()

    If Trim(cmbLocal.Text) = "CMCE" Or Trim(cmbLocal.Text) = "MC85E" And LocalizaTipoNatureza(Val(txtCodigoOperacao.Text) + 1000 + Val(cmbNaturezaOperacao.Text)) = "RCt" Then
        MsgBox "Esta nota não precisa conter vencimentos!", vbExclamation, "Atenção"
    
        Exit Sub
    End If
    
'    frmPagtoFornecedor.txtFornecedor.Text = Mid(txtNomeFantasia.Text, 1, 3)
'    frmPagtoFornecedor.txtSerie.Text = txtSerie.Text
'    frmPagtoFornecedor.TxtNota.Text = txtNotaFiscal.Text
'    frmPagtoFornecedor.txtDataEmissao.Text = mskDataEmissao.Text
'    frmPagtoFornecedor.txtDataRecebimento.Text = mskDataRecebimento.Text
    
    frmPagtoFornecedor.Show 1

End Sub

Private Sub Command1_Click()
Call LerItemCritica
End Sub

Private Sub Form_Load()
   frmEntradaNotaFiscal2.top = (Screen.Height - frmEntradaNotaFiscal2.Height) / 2
   frmEntradaNotaFiscal2.left = (Screen.Width - frmEntradaNotaFiscal2.Width) / 2
        
    Screen.MousePointer = 11
    
  '  Status "Inicializando a Tela de Entrada de Notas Fiscais..."
    
    Pesquisou = False
    
    GetAsyncKeyState vbKeyTab
    InicializaClasses
    MostraData
    MontaCombos
    ResetMe
    lblControlaTela.Caption = ""
    
    'Status "Pronto."
    
    Screen.MousePointer = 0

End Sub

Private Sub Form_MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)

    pnlObsPedido.Visible = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
'OutraTela = "N"
End Sub

Private Sub Frame1_MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)

    pnlObsPedido.Visible = False

End Sub

Private Sub grdItens_Click()

    lblPedido.Visible = True
    lblComprador.Visible = True
    lblCFOP.Visible = True
    lblFrete.Visible = True
    lblEntrega.Visible = True
    lblEntrega2.Visible = True
    lblCondPagto.Visible = True
    lblCondPagto2.Visible = True

    PreencheCapa

End Sub

Private Sub grdItens_DblClick()

    On Error Resume Next

    DesceDados
    PreencheCapa
    txtQuantidade.SetFocus
    
    Err.Clear

End Sub

Sub DesceDados()

    Dim rdoPedido As New ADODB.Recordset
    
    If Not SoExclusao Then
        If grdItens.Row > 0 And grdItens.TextMatrix(1, 0) <> "" And grdItens.RowIsVisible(grdItens.Row) Then
            If Trim(grdItens.TextMatrix(grdItens.Row, 0)) <> RefDevolucao Then
                
                SQL = "Select PI_PrecoUnitario, PI_PercentualDesconto, " _
                    & "PI_AliquotaIPI " _
                    & "From ItemPedido " _
                    & "Where PI_NumeroPedido = " & Trim(grdItens.TextMatrix(grdItens.Row, 1)) _
                    & " and PI_Referencia = '" & Trim(grdItens.TextMatrix(grdItens.Row, 0)) & "'"
                    
                Set rdoPedido = rdoCnSupBatch.OpenResultset(SQL, Options:=rdExecDirect)
                
                txtPrecoUnitario.Text = ""
                txtAliquotaIPI.Text = "0,00"
                txtPercentualDesconto.Text = "0,00"
                
                If Not rdoPedido.EOF Then
                   txtPrecoUnitario.Text = Format(rdoPedido("PI_PrecoUnitario") - ((rdoPedido("PI_PrecoUnitario") * rdoPedido("PI_PercentualDesconto")) / 100) + ((rdoPedido("PI_PrecoUnitario") * rdoPedido("PI_AliquotaIPI")) / 100), "###,###,###,##0.00")
                   txtAliquotaIPI.Text = rdoPedido("PI_AliquotaIPI")
                End If
                   
'                If Trim(grdItens.TextMatrix(grdItens.Row, 6)) <> "" Then
'                    txtPercentualDesconto.Text = Trim(grdItens.TextMatrix(grdItens.Row, 6))
'                End If
                
                txtReferencia.Text = Trim(grdItens.TextMatrix(grdItens.Row, 0))
                txtNumeroPedido.Text = Trim(grdItens.TextMatrix(grdItens.Row, 1))
                txtQuantidade.Text = Trim(grdItens.TextMatrix(grdItens.Row, 3))
                
                If Trim(grdItens.TextMatrix(grdItens.Row, 3)) = "" Then
                    clsGravaItem.Adicionar = True
                Else
                    clsGravaItem.Adicionar = False
                End If
            End If
        End If
    End If

End Sub

Private Sub grdItens_KeyDown(KeyCode As Integer, shift As Integer)

    Dim rdoExclui As New ADODB.Recordset
    
    On Error Resume Next
    
    If KeyCode = vbKeyDelete And grdItens.RowIsVisible(grdItens.Row) And Trim(grdItens.TextMatrix(grdItens.Row, 4)) <> "" Then
        If MsgBox("Confirma a exclusão do item selecionado?", vbQuestion + vbYesNo + vbDefaultButton2, "Exclusão de Item") = vbYes Then
            Screen.MousePointer = 11
            
            Err.Clear
            
            Set rdoExclui = rdoCnSupBatch.OpenResultset("ExcluiItemNFCompra '" & grdItens.TextMatrix(grdItens.Row, 0) & "', " & grdItens.TextMatrix(grdItens.Row, 1) & ", " & txtNotaFiscal.Text & ", '" & txtSerie.Text & "', " & Val(txtNomeFantasia.Text) & ", " & IIf(grdItens.RowData(grdItens.Row) < 255, grdItens.RowData(grdItens.Row), 0), Options:=rdExecDirect)
            
            If Err.Number = 0 Then
                If grdItens.TextMatrix(grdItens.Row, 0) <> RefDevolucao Then
                    grdItens.TextMatrix(grdItens.Row, 3) = ""
                    grdItens.TextMatrix(grdItens.Row, 4) = ""
                    'grdItens.TextMatrix(grdItens.Row, 5) = ""
                    'grdItens.TextMatrix(grdItens.Row, 6) = ""
                    grdItens.TextMatrix(grdItens.Row, 7) = ""
                    grdItens.TextMatrix(grdItens.Row, 9) = ""
                    grdItens.FillStyle = flexFillRepeat
                    grdItens.CellForeColor = vbBlack
                    grdItens.FillStyle = flexFillSingle

                Else
                    If grdItens.Rows = 2 Then
                        clsItensNota.Clear
                    Else
                        grdItens.RemoveItem grdItens.Row
                        
                        grdItens.Row = 1
                        grdItens.Col = 0
                        grdItens.ColSel = grdItens.Cols - 1
                    End If
                End If
                
                txtTotalCalculado.Text = Format(rdoExclui("ValorCalculado"), "###,###0.00")
                CalculaBateNota
                rdoExclui.Close
            Else
                rdoExclui.Close
                MsgBox "Erro excluindo item de Nota Fiscal. Tente novamente.", vbCritical, "Erro"
                MostraErro
            End If
            Screen.MousePointer = 0
        End If
    End If

End Sub

Private Sub grdItens_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        DesceDados
        PreencheCapa
        txtQuantidade.SetFocus
    End If

End Sub
Function PreencheCapa()
lblPedido.Caption = "Pedido" & " " & grdItens.TextMatrix(grdItens.Row, 1)
lblComprador.Caption = "Comprador" & " " & grdItens.TextMatrix(grdItens.Row, 10)
lblCFOP.Caption = "CFOP" & " " & grdItens.TextMatrix(grdItens.Row, 14)
lblFrete.Caption = "Frete" & " " & grdItens.TextMatrix(grdItens.Row, 13)
lblEntrega.Caption = "Semana de Entrega"
lblEntrega2.Caption = Format(grdItens.TextMatrix(grdItens.Row, 11), "##/####")
lblCondPagto.Caption = "Condição Pagamento"
lblCondPagto2.Caption = Format(grdItens.TextMatrix(grdItens.Row, 12), "000/000/000/000/000")
End Function
Private Sub GrdItens_LeaveCell()

    clsGravaItem.Limpar
    PreencheCapa
    DesceDados

End Sub

Private Sub grdItens_MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)

    If grdItens.TextMatrix(1, 0) <> "" And grdItens.MouseRow > 0 And grdItens.MouseCol = 1 Then
        MostraObs Val(grdItens.TextMatrix(grdItens.MouseRow, 1))
    Else
        pnlObsPedido.Visible = False
    End If

End Sub

Sub MostraObs(ByVal NumeroPedido As Long)

    Dim Indice As Long
    Dim StringObs As String
    Dim Maximo As Long
    
    Maximo = UBound(Observacoes)
    
    StringObs = ""
    For Indice = 1 To Maximo Step 1
        If Observacoes(Indice).NumeroPedido = NumeroPedido Then
            StringObs = Trim(Observacoes(Indice).Obs)
            If StringObs = "" Then
                StringObs = "Nenhuma observação."
            End If
            Exit For
        End If
    Next Indice
    
    pnlObsPedido.Caption = StringObs
    If StringObs <> "" Then
        pnlObsPedido.Visible = True
    Else
        pnlObsPedido.Visible = False
    End If

End Sub

Private Sub grdItens_RowColChange()

    DesceDados
    pnlObsPedido.Visible = False

End Sub

Private Sub Label37_Click()

End Sub

Private Sub mskDataEmissao_GotFocus()
    
    mskDataEmissao.SelStart = 0
    mskDataEmissao.SelLength = Len(mskDataEmissao.Text)

End Sub

Private Sub pnlItens_MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)

    pnlObsPedido.Visible = False

End Sub

Private Sub txtAliquotaIPI_GotFocus()
    
    txtAliquotaIPI.SelStart = 0
    txtAliquotaIPI.SelLength = Len(txtAliquotaIPI.Text)

End Sub

Private Sub txtAliquotaIcms_GotFocus()
    
    txtaliquotaicms.SelStart = 0
    txtaliquotaicms.SelLength = Len(txtaliquotaicms.Text)

End Sub

Private Sub txtAliquotaIPI_KeyDown(KeyCode As Integer, shift As Integer)

    Dim Linha As Long

    Linha = grdItens.Row
    
    clsItensNota.SelecionaLinha KeyCode
    
    If grdItens.Row <> Linha Then
        DesceDados
        PreencheCapa
    End If

End Sub

Private Sub txtAliquotaIPI_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If Not txtCGC.Enabled Then
            If txtReferencia.Text = RefDevolucao Then
                txtNumeroPedido.Text = "0"
            End If
            clsGravaItem.Verificar
        End If
    Else
        VerteclaVirgula Me.ActiveControl, KeyAscii
    End If

End Sub

Private Sub txtAliquotaIPI_LostFocus()

    txtAliquotaIPI.Text = Format(txtAliquotaIPI.Text, "0.00")

End Sub

Private Sub txtAliquotaICMS_LostFocus()

    txtaliquotaicms.Text = Format(txtaliquotaicms.Text, "0.00")

End Sub

Private Sub txtAliquotaICMS_KeyPress(KeyAscii As Integer)

    VerteclaVirgula Me.ActiveControl, KeyAscii

End Sub

Private Sub txtBaseICMS_GotFocus()
    
    txtBaseICMS.SelStart = 0
    txtBaseICMS.SelLength = Len(txtBaseICMS.Text)

End Sub

Private Sub txtBaseICMS_KeyPress(KeyAscii As Integer)

    VerteclaVirgula Me.ActiveControl, KeyAscii

End Sub

Private Sub txtBaseICMS_LostFocus()

    txtBaseICMS.Text = Format(txtBaseICMS.Text, "###,###,###,###0.00")

End Sub

Private Sub txtCGC_GotFocus()
    
    txtCGC.SelStart = 0
    txtCGC.SelLength = Len(txtCGC.Text)

End Sub

Private Sub txtCGC_KeyPress(KeyAscii As Integer)

    VerTecla KeyAscii

End Sub

Private Sub txtCGC_LostFocus()
    
    Dim rdoFornecedor As New ADODB.Recordset
    
    If GetAsyncKeyState(vbKeyTab) <> 0 Then
        If Trim(txtCGC.Text) <> "" Then
            Screen.MousePointer = 11
            
            SQL = "select FO_CodigoFornecedor, FO_NomeFantasia from Fornecedor where FO_CGC = '" & txtCGC.Text & "'"
            rdoFornecedor.CursorLocation = adUseClient
            rdoFornecedor.Open SQL, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
            
            'Set rdoFornecedor = rdoCnSup.OpenResultset("select FO_CodigoFornecedor, FO_NomeFantasia from Fornecedor where FO_CGC = '" & txtCGC.Text & "'", Options:=rdExecDirect)
            Screen.MousePointer = 0
            If rdoFornecedor.EOF Then
                rdoFornecedor.Close
                MsgBox "Fornecedor não cadastrado!", vbExclamation, "Atenção"
                txtNomeFantasia.Text = ""
                txtCGC.Enabled = True
                txtCGC.SetFocus
                Exit Sub
            End If
            txtNomeFantasia.Text = Format(rdoFornecedor("FO_CodigoFornecedor"), "000") & " - " & rdoFornecedor("FO_NomeFantasia")
            txtCGC.Enabled = False
            rdoFornecedor.Close
        Else
            txtNomeFantasia.Text = ""
        End If
    ElseIf Trim(txtCGC.Text) = "" Then
        txtNomeFantasia.Text = ""
    End If

End Sub

Private Sub txtCodigoOperacao_GotFocus()
    
    txtCodigoOperacao.SelStart = 0
    txtCodigoOperacao.SelLength = Len(txtCodigoOperacao.Text)

End Sub

Private Sub txtCodigoOperacao_KeyPress(KeyAscii As Integer)

    VerTecla KeyAscii

End Sub

Private Sub txtCodigoOperacao_LostFocus()

    Dim SQL As String
    Dim RdoCodOper As New ADODB.Recordset
    
    If Val(txtCodigoOperacao.Text) = 0 Or Trim(txtCodigoOperacao.Text) = "" Then
             MsgBox "Informe o Código de Operação.", vbExclamation, "Atenção"
             'txtNotaFiscal.SetFocus
             Exit Sub
    End If
    SQL = "Select CN_CodigoOperacaoAntigo from CodigoOperacaoNovo where CN_CodigoOperacaoNovo=" & Mid(txtCodigoOperacao, 1, 4) & ""
    RdoCodOper.CursorLocation = adUseClient
    RdoCodOper.Open SQL, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
    
    'Set RdoCodOper = rdoCnSup.OpenResultset(SQL)
    
    If Not RdoCodOper.EOF Then
       WCodigoOperacaoVelho = RdoCodOper("CN_CODIGOOPERACAOANTIGO")
    Else
       WCodigoOperacaoVelho = txtCodigoOperacao.Text
    End If
   
    MontaComboNatureza Val(WCodigoOperacaoVelho), cmbNaturezaOperacao
    
    If cmbNaturezaOperacao.ListCount > 0 Then
        cmbNaturezaOperacao.ListIndex = 0
    End If

End Sub

Private Sub txtDespesas_Change()

    If Pesquisou Then
        If IsNumeric(Format(txtDespesas.Text, "0.00")) Then
            txtTotalCalculado.Text = Format(TotalCalculado + CDbl(Format(txtFrete.Text, "0.00")) + CDbl(Format(txtEmbalagem.Text, "0.00")) + Despesas + CDbl(Format(txtJuros.Text, "0.00")) + CDbl(Format(txtOutros.Text, "0.00")) + (CDbl(Format(txtDespesas.Text, "0.00")) - Despesas), "###,###,###0.00")
        Else
            txtTotalCalculado.Text = Format(TotalCalculado + CDbl(Format(txtFrete.Text, "0.00")) + CDbl(Format(txtFrete.Text, "0.00")) + CDbl(Format(txtFrete.Text, "0.00")) + CDbl(Format(txtFrete.Text, "0.00")), "###,###,###0.00")
        End If
        
        BateNota = CDbl(Format(txtTotalCalculado.Text, "0.00")) - CDbl(Format(txtValorTotalNota.Text, "0.00"))
        txtBateNota.Text = Format(BateNota, "###,###,###0.00")
    End If
 
End Sub

Private Sub txtDespesas_GotFocus()
    
    txtDespesas.SelStart = 0
    txtDespesas.SelLength = Len(txtDespesas.Text)

End Sub

Private Sub txtDespesas_KeyPress(KeyAscii As Integer)

    VerteclaVirgula Me.ActiveControl, KeyAscii

End Sub

Private Sub txtDespesas_LostFocus()

    txtDespesas.Text = Format(txtDespesas.Text, "###,###,###,###0.00")
    
    If Numeros(txtDespesas.Text) = "" Then
        txtDespesas.Text = "0,00"
    End If

End Sub

Private Sub txtEmbalagem_Change()

    If Pesquisou Then
        If IsNumeric(Format(txtEmbalagem.Text, "0.00")) Then
            txtTotalCalculado.Text = Format(TotalCalculado + CDbl(Format(txtFrete.Text, "0.00")) + Embalagem + CDbl(Format(txtDespesas.Text, "0.00")) + CDbl(Format(txtJuros.Text, "0.00")) + CDbl(Format(txtOutros.Text, "0.00")) + (CDbl(Format(txtEmbalagem.Text, "0.00")) - Embalagem), "###,###,###0.00")
        Else
            txtTotalCalculado.Text = Format(TotalCalculado + CDbl(Format(txtFrete.Text, "0.00")) + CDbl(Format(txtDespesas.Text, "0.00")) + CDbl(Format(txtJuros.Text, "0.00")) + CDbl(Format(txtOutros.Text, "0.00")), "###,###,###0.00")
        End If
        
        BateNota = CDbl(Format(txtTotalCalculado.Text, "0.00")) - CDbl(Format(txtValorTotalNota.Text, "0.00"))
        txtBateNota.Text = Format(BateNota, "###,###,###0.00")
    End If
 
 End Sub

Private Sub txtEmbalagem_GotFocus()
    
    txtEmbalagem.SelStart = 0
    txtEmbalagem.SelLength = Len(txtEmbalagem.Text)

End Sub

Private Sub txtEmbalagem_KeyPress(KeyAscii As Integer)

    VerteclaVirgula Me.ActiveControl, KeyAscii

End Sub

Private Sub txtEmbalagem_LostFocus()

    txtEmbalagem.Text = Format(txtEmbalagem.Text, "###,###,###,###0.00")

    If Numeros(txtEmbalagem.Text) = "" Then
        txtEmbalagem.Text = "0,00"
    End If

End Sub

Private Sub txtFrete_Change()

    If Pesquisou Then
        If IsNumeric(Format(txtFrete.Text, "0.00")) Then
            txtTotalCalculado.Text = Format(TotalCalculado + Frete + CDbl(Format(txtEmbalagem.Text, "0.00")) + CDbl(Format(txtDespesas.Text, "0.00")) + CDbl(Format(txtJuros.Text, "0.00")) + CDbl(Format(txtOutros.Text, "0.00")) + (CDbl(Format(txtFrete.Text, "0.00")) - Frete), "###,###,###0.00")
        Else
            txtTotalCalculado.Text = Format(TotalCalculado + CDbl(Format(txtEmbalagem.Text, "0.00")) + CDbl(Format(txtDespesas.Text, "0.00")) + CDbl(Format(txtJuros.Text, "0.00")) + CDbl(Format(txtOutros.Text, "0.00")), "###,###,###0.00")
        End If
        
        BateNota = CDbl(Format(txtTotalCalculado.Text, "0.00")) - CDbl(Format(txtValorTotalNota.Text, "0.00"))
        txtBateNota.Text = Format(BateNota, "###,###,###0.00")
    End If
    
End Sub

Private Sub txtFrete_GotFocus()
    
    txtFrete.SelStart = 0
    txtFrete.SelLength = Len(txtFrete.Text)

End Sub

Private Sub txtFrete_KeyPress(KeyAscii As Integer)

    VerteclaVirgula Me.ActiveControl, KeyAscii

End Sub

Private Sub txtFrete_LostFocus()

    txtFrete.Text = Format(txtFrete.Text, "###,###,###,###0.00")
    
    If Numeros(txtFrete.Text) = "" Then
        txtFrete.Text = "0,00"
    End If

End Sub

Private Sub txtJuros_Change()

    If Pesquisou Then
        If IsNumeric(Format(txtJuros.Text, "0.00")) Then
            txtTotalCalculado.Text = Format(TotalCalculado + CDbl(Format(txtFrete.Text, "0.00")) + CDbl(Format(txtEmbalagem.Text, "0.00")) + CDbl(Format(txtDespesas.Text, "0.00")) + Juros + CDbl(Format(txtOutros.Text, "0.00")) + (CDbl(Format(txtJuros.Text, "0.00")) - Juros), "###,###,###0.00")
        Else
            txtTotalCalculado.Text = Format(TotalCalculado + CDbl(Format(txtFrete.Text, "0.00")) + CDbl(Format(txtEmbalagem.Text, "0.00")) + CDbl(Format(txtDespesas.Text, "0.00")) + CDbl(Format(txtOutros.Text, "0.00")), "###,###,###0.00")
        End If
        
        BateNota = CDbl(Format(txtTotalCalculado.Text, "0.00")) - CDbl(Format(txtValorTotalNota.Text, "0.00"))
        txtBateNota.Text = Format(BateNota, "###,###,###0.00")
    End If
 
End Sub

Private Sub txtJuros_GotFocus()
    
    txtJuros.SelStart = 0
    txtJuros.SelLength = Len(txtJuros.Text)

End Sub

Private Sub txtJuros_KeyPress(KeyAscii As Integer)

    VerteclaVirgula Me.ActiveControl, KeyAscii

End Sub

Private Sub txtJuros_LostFocus()

    txtJuros.Text = Format(txtJuros.Text, "###,###,###,###0.00")

    If Numeros(txtJuros.Text) = "" Then
        txtJuros.Text = "0,00"
    End If

End Sub

Private Sub txtNotaFiscal_GotFocus()
    
    txtNotaFiscal.SelStart = 0
    txtNotaFiscal.SelLength = Len(txtNotaFiscal.Text)

End Sub

Private Sub txtNotaFiscal_KeyPress(KeyAscii As Integer)

    VerTecla KeyAscii

End Sub

Private Sub txtOutros_Change()

    If Pesquisou Then
        If IsNumeric(Format(txtOutros.Text, "0.00")) Then
            txtTotalCalculado.Text = Format(TotalCalculado + CDbl(Format(txtFrete.Text, "0.00")) + CDbl(Format(txtEmbalagem.Text, "0.00")) + CDbl(Format(txtDespesas.Text, "0.00")) + CDbl(Format(txtJuros.Text, "0.00")) + Outros + (CDbl(Format(txtOutros.Text, "0.00")) - Outros), "###,###,###0.00")
        Else
            txtTotalCalculado.Text = Format(TotalCalculado + CDbl(Format(txtFrete.Text, "0.00")) + CDbl(Format(txtEmbalagem.Text, "0.00")) + CDbl(Format(txtDespesas.Text, "0.00")) + CDbl(Format(txtJuros.Text, "0.00")), "###,###,###0.00")
        End If
        
        BateNota = CDbl(Format(txtTotalCalculado.Text, "0.00")) - CDbl(Format(txtValorTotalNota.Text, "0.00"))
        txtBateNota.Text = Format(BateNota, "###,###,###0.00")
    End If
 
End Sub

Private Sub txtOutros_GotFocus()
    
    txtOutros.SelStart = 0
    txtOutros.SelLength = Len(txtOutros.Text)

End Sub

Private Sub txtOutros_KeyPress(KeyAscii As Integer)

    VerteclaVirgula Me.ActiveControl, KeyAscii

End Sub

Private Sub txtOutros_LostFocus()

    txtOutros.Text = Format(txtOutros.Text, "###,###,###,###0.00")
    
    If Numeros(txtOutros.Text) = "" Then
        txtOutros.Text = "0,00"
    End If

End Sub

Private Sub txtPercentualDesconto_GotFocus()
    
    txtPercentualDesconto.SelStart = 0
    txtPercentualDesconto.SelLength = Len(txtPercentualDesconto.Text)

End Sub

Private Sub txtPercentualDesconto_KeyDown(KeyCode As Integer, shift As Integer)

    Dim Linha As Long

    Linha = grdItens.Row
    
    clsItensNota.SelecionaLinha KeyCode
    
    If grdItens.Row <> Linha Then
        DesceDados
        PreencheCapa
        txtQuantidade.SetFocus
    End If

End Sub

Private Sub txtPercentualDesconto_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If Not txtCGC.Enabled Then
            If txtReferencia.Text = RefDevolucao Then
                txtNumeroPedido.Text = "0"
            End If
            clsGravaItem.Verificar
        End If
    Else
        VerteclaVirgula Me.ActiveControl, KeyAscii
    End If

End Sub

Private Sub txtPercentualDesconto_LostFocus()

    txtPercentualDesconto.Text = Format(txtPercentualDesconto.Text, "0.00")
    
    If GetAsyncKeyState(vbKeyTab) <> 0 And Me.ActiveControl.Name <> txtAliquotaIPI.Name And txtReferencia.Text <> "" Then
        txtPercentualDesconto_KeyPress vbKeyReturn
    End If

    txtQuantidade.SetFocus

End Sub
Private Sub LerItemCritica()

Dim rsCritica As New ADODB.Recordset

    SQL = ""
    SQL = "Exec CriticaNotaFiscalCompra " & txtNotaFiscal.Text & ", '" & txtSerie.Text & "', " _
        & "" & Val(txtNomeFantasia.Text) & ", '" & grdItens.TextMatrix(grdItens.Row, 0) & "'"
    Set rsCritica = rdoCnSupBatch.OpenResultset(SQL)
    
    If Not rsCritica.EOF Then
        If Trim(rsCritica("Critica")) <> "" Then
            grdItens.RowSel = grdItens.Row
            grdItens.FillStyle = flexFillRepeat
            grdItens.CellForeColor = &HDD&
            grdItens.FillStyle = flexFillSingle
            grdItens.TextMatrix(grdItens.Row, 9) = rsCritica("Critica")
        End If
    End If

    rsCritica.Close


End Sub
Function DistribuicaoAutomatica()
    Dim VarF As Integer

    If Trim(cmbLocal.Text) <> "CD" Then
    
        Exit Function
    
    End If
    
    Screen.MousePointer = 11
    
    Do While Not adoChkFornecedor.EOF
        If Mid(grdItens.TextMatrix(I, 0), 1, 3) = adoChkFornecedor("Fornecedor") Then
            TotalGrade = 0
            QtdeDistribuicao = 0
            PercCalculado = 0
            
            Call VerificaGrade
            
            wWhere = " lo_loja not in ('182','183','184','185','CD','CMC','CMCE','CMCS','CONSO') "
            
            Call VerificaLoja
            
            If rsDados.EOF = True Then
            
                Do While Not rsLoja.EOF
                
                    SQL = ""
                    
                    SQL = "INSERT INTO DistribuicaoAutomatica(DA_NotaFiscal,DA_Serie,DA_LojaDestino," & _
                        "DA_Referencia,DA_Situacao,DA_CodigoFornecedor,DA_PercentualDistribuicao,DA_DistribuicaoAtendida)" & _
                        " Values(" & txtNotaFiscal.Text & ",'" & txtSerie.Text & "','" & Trim(rsLoja("LO_Loja")) & "','" & _
                        grdItens.TextMatrix(I, 0) & " ','A','" & Fornecedor & "',0,0)"
                    
                    rdoCnSupBatch.Execute (SQL)
                    
                    rsLoja.MoveNext
                
                Loop
            
            End If
            Call ExecutaCalculos
        End If
        adoChkFornecedor.MoveNext
    Loop
    
    Screen.MousePointer = 0
    
End Function
Function ExecutaCalculos()

    Dim PrimeiraVez As Integer
    
    PrimeiraVez = 1
    
    Residencia = grdItens.TextMatrix(I, 15) 'RESIDENCIA
    QtdeCalculada = grdItens.TextMatrix(I, 3) 'QUANTIDADE A SER DISTRIBUIDA
    
    If Residencia = 1 Or Residencia = 99 Then
    
        wWhere = " lo_loja not in ('182','183','184','185','CD','CMC','CMCE','CMCS','CONSO') "
    
    Else
    
        wWhere = " LO_OrdemLoja = '" & Residencia & "' "
    
    End If
    
    If Residencia = 1 Then
    
        Exit Function
    
    End If
    
    Do While QtdeCalculada > 0
    
        wFlagAtendimento = ""
        
        Call VerificaLoja
        
        Do While Not rsLoja.EOF
        
            Call VerificaEstoque
            
            If rsPerc.EOF = False Then
            
                Call VerificaAtendimento 'Verifica SE todas as Lojas excederam
                
                If rsFaltaLoja("FaltaLoja") > 0 Then 'SE não excedeu perguntar se a soma da distribuicao atendida + estoque + romaneio + transito for maior que o maximo
                
                    If (rsPerc("DA_DistribuicaoAtendida") + rsPerc("ES_Estoque") + rsPerc("ES_Romaneio") + rsPerc("ES_Transito")) < rsPerc("ES_EstoqueMaximo") + rsPerc("ES_Display") Then
                        
                        Call VerificaEstoque
                        
                        Call ProcedimentoDistribuicao
                        
                    End If
                
                ElseIf QtdeCalculada >= 0 Then 'SENÃO SE a quantidade não foi totalmente distribuida e mesmo assim TODAS as lojas já estão totalmente atendidas
                
                    If PrimeiraVez = 1 Then
                    
                        Call VerificaLoja
                        
                        PrimeiraVez = 0
                        
                    End If
                    
                    Call ProcedimentoDistribuicao 'Continuar a rotina de distribuição mesmo excedendo o maximo
                    
                End If
                
                If QtdeCalculada <= 0 Then
                
                    Exit Do
                
                End If
            
            End If
            
            rsLoja.MoveNext
        
        Loop
        
        
        
        If wFlagAtendimento = "" Then
        
            Exit Do
        
        End If
        
    Loop

End Function

Function VerificaEstoque()
    
    SQL = ""
    
    SQL = "SELECT ES_Display,CI_Item,DA_DistribuicaoAtendida,ES_EstoqueMinimo,ES_Estoque,ES_EstoqueMaximo,ES_Romaneio,ES_Transito,CI_Quantidade,PR_Residencia " & _
          "FROM ItemNFCompra,DistribuicaoAutomatica,Estoque,Produto WHERE (ES_EstoqueMinimo + ES_Display) > 0 " & _
          " and PR_Referencia = ES_Referencia and DA_Referencia = ES_Referencia and DA_NotaFiscal = CI_NotaFiscal " & _
          " and DA_LojaDestino = ES_Loja " & _
          " and CI_NotaFiscal = '" & txtNotaFiscal & _
          "' and ES_Referencia = '" & grdItens.TextMatrix(I, 0) & "' " & _
          " and ES_Loja = '" & Trim(rsLoja("LO_Loja")) & "' " & _
          " and CI_Referencia = ES_Referencia"
    
    Set rsPerc = rdoCnSupBatch.OpenResultset(SQL)

End Function

Function ProcedimentoDistribuicao()

    If (rsPerc("ES_Estoque") + rsPerc("ES_Romaneio") + rsPerc("ES_Transito") + rsPerc("DA_DistribuicaoAtendida")) > rsPerc("ES_EstoqueMaximo") + rsPerc("ES_Display") Then

         If rsPerc("ES_EstoqueMaximo") + rsPerc("ES_Display") < QtdeCalculada Then

            QtdeCalculo = rsPerc("ES_EstoqueMaximo") + rsPerc("ES_Display")

         Else

            QtdeCalculo = QtdeCalculada

         End If

    ElseIf rsPerc("ES_EstoqueMinimo") + rsPerc("ES_Display") > QtdeCalculada Then

        QtdeCalculo = QtdeCalculada

    Else

        QtdeCalculo = rsPerc("ES_EstoqueMinimo") + rsPerc("ES_Display")

    End If

        

    SQL = ""
    SQL = "SELECT ES_Estoque,ES_Romaneio,ES_Transito FROM ESTOQUE WHERE ES_Loja = '" & Trim(rsLoja("LO_Loja")) & "' " & _
          " and ES_Referencia = '" & grdItens.TextMatrix(I, 0) & "'"
    Set adoConsEstoque = rdoCnSup.OpenResultset(SQL)
    
    SQL = "UPDATE DistribuicaoAutomatica SET " & _
        " DA_Estoque = '" & adoConsEstoque("ES_Estoque") & "' " & _
        ",DA_Romaneio = '" & adoConsEstoque("ES_Romaneio") & "' " & _
        ",DA_Transito = '" & adoConsEstoque("ES_Transito") & "' " & _
        ",DA_DistribuicaoAtendida = DA_DistribuicaoAtendida + " & ConverteVirgula(QtdeCalculo) & _
        " WHERE DA_NotaFiscal = " & txtNotaFiscal.Text & _
        " and DA_Serie = '" & txtSerie.Text & _
        "' and DA_Referencia = '" & grdItens.TextMatrix(I, 0) & "' " & _
        " and DA_LojaDestino = '" & Trim(rsLoja("LO_Loja")) & "'"
    
    rdoCnSupBatch.Execute (SQL)
    
    QtdeCalculada = QtdeCalculada - QtdeCalculo
    
    'SUBTRAI O ESTOQUE MINIMO DA QUANTIDADE CALCULADA
    
    wFlagAtendimento = "*"
    
    'Call VerificaAtendimento

End Function

Function VerificaAtendimento()
    
    SQL = "SELECT isnull(count(es_estoque),0) as FaltaLoja FROM Estoque,DistribuicaoAutomatica,ItemNfCompra,capanfcompra  WHERE " & _
          "DA_Notafiscal = CI_NotaFiscal and DA_Referencia = ES_Referencia and da_lojadestino = es_loja and " & _
          "CI_NotaFiscal = " & txtNotaFiscal.Text & " and DA_Referencia = '" & grdItens.TextMatrix(I, 0) & "' and DA_Referencia = ES_Referencia and DA_CodigoFornecedor = CC_Fornecedor and CC_NotaFiscal = CI_NotaFiscal and CC_Serie = CI_Serie and DA_Notafiscal = CI_NotaFiscal and DA_Serie = CI_Serie and ES_Loja = DA_LojaDestino " & _
          " and (ES_Estoque + ES_Transito + ES_Romaneio + DA_DistribuicaoAtendida) < (ES_EstoqueMaximo + ES_Display) "

    Set rsFaltaLoja = rdoCnSupBatch.OpenResultset(SQL)

End Function

Function VerificaGrade()
    
    SQL = ""

    SQL = "SELECT * FROM DistribuicaoAutomatica WHERE DA_NotaFiscal = " & txtNotaFiscal.Text & _
    " and DA_Referencia = '" & grdItens.TextMatrix(I, 0) & "'"

    Set rsDados = rdoCnSup.OpenResultset(SQL)

End Function

Function VerificaLoja()

    SQL = ""

    SQL = "SELECT LO_Loja FROM Loja,Estoque Where LO_Situacao = 'A' and LO_Loja = ES_Loja " & _
          " and ES_Referencia = '" & grdItens.TextMatrix(I, 0) & "' and " & _
            wWhere & _
          " Order By ES_EstoqueMaximo Desc"

    Set rsLoja = rdoCnSup.OpenResultset(SQL)

End Function

Private Sub txtPesquisaDescricao_LostFocus()
    
    If grdItens.TextMatrix(1, 0) <> "" And txtPesquisaDescricao.Text <> "" Then
        PintaGrid 2
    End If

End Sub

Private Sub txtPesquisaPedido_LostFocus()
    
    If grdItens.TextMatrix(1, 0) <> "" And Val(txtPesquisaPedido.Text) > 0 Then
        PintaGrid 1
    End If

End Sub

Sub PintaGrid(ByVal TipoPesquisa As Long)

    Dim Maximo As Long
    Dim Linha As Long
    Dim Procurado As String
    Dim Colunas As Long
    Dim PrimeiraLinha As Long

    Maximo = grdItens.Rows - 1
    Colunas = grdItens.Cols - 1
    
    grdItens.Redraw = False
    grdItens.FillStyle = flexFillRepeat
    
    grdItens.Row = 1
    grdItens.Col = 0
    grdItens.RowSel = Maximo
    grdItens.ColSel = Colunas
    grdItens.CellBackColor = 0
    grdItens.RowSel = 1
    
    PrimeiraLinha = 0
    
    If TipoPesquisa = 1 Then 'Pedido
        Procurado = Val(txtPesquisaPedido.Text)
        For Linha = 1 To Maximo Step 1
            If grdItens.TextMatrix(Linha, 1) = Procurado Then
                If PrimeiraLinha = 0 Then
                    PrimeiraLinha = Linha
                End If
                grdItens.Row = Linha
                grdItens.ColSel = Colunas
                grdItens.CellBackColor = AmareloGrid
            End If
        Next Linha
        txtPesquisaPedido.Text = ""
    Else    'Descrição
        Procurado = txtPesquisaDescricao.Text
        For Linha = 1 To Maximo Step 1
            If InStr(1, grdItens.TextMatrix(Linha, 2), Procurado, vbTextCompare) <> 0 Then
                If PrimeiraLinha = 0 Then
                    PrimeiraLinha = Linha
                End If
                grdItens.Row = Linha
                grdItens.ColSel = Colunas
                grdItens.CellBackColor = AmareloGrid
            End If
        Next Linha
        txtPesquisaDescricao.Text = ""
    End If
    
    grdItens.FillStyle = flexFillSingle
    grdItens.Redraw = True
    
    If PrimeiraLinha > 0 Then
        grdItens.Row = PrimeiraLinha
        clsItensNota.TornaLinhaVisivel PrimeiraLinha
    End If
    grdItens.ColSel = Colunas

End Sub

Private Sub txtPrecoUnitario_GotFocus()
    
    txtPrecoUnitario.SelStart = 0
    txtPrecoUnitario.SelLength = Len(txtPrecoUnitario.Text)

End Sub

Private Sub txtPrecoUnitario_KeyDown(KeyCode As Integer, shift As Integer)

    Dim Linha As Long

    Linha = grdItens.Row
    
    clsItensNota.SelecionaLinha KeyCode
    
    If grdItens.Row <> Linha Then
        PreencheCapa
        DesceDados
    End If

End Sub

Private Sub txtPrecoUnitario_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If Not txtCGC.Enabled Then
            If txtReferencia.Text = RefDevolucao Then
                txtNumeroPedido.Text = "0"
            End If
            clsGravaItem.Verificar
        End If
    Else
        VerteclaVirgula Me.ActiveControl, KeyAscii
    End If

End Sub

Private Sub txtPrecoUnitario_LostFocus()

    txtPrecoUnitario.Text = Format(txtPrecoUnitario.Text, "###,###,###,###0.00")

End Sub

Private Sub txtQuantidade_GotFocus()
    
    txtQuantidade.SelStart = 0
    txtQuantidade.SelLength = Len(txtQuantidade.Text)

End Sub

Private Sub txtQuantidade_KeyDown(KeyCode As Integer, shift As Integer)

    Dim Linha As Long

    Linha = grdItens.Row
    
    clsItensNota.SelecionaLinha KeyCode
    
    If grdItens.Row <> Linha Then
        DesceDados
        PreencheCapa
    End If

End Sub

Private Sub txtQuantidade_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If Not txtCGC.Enabled Then
            If txtReferencia.Text = RefDevolucao Then
                txtNumeroPedido.Text = "0"
            End If
            clsGravaItem.Verificar
        End If
    Else
        VerTecla KeyAscii
    End If

End Sub

Private Sub txtQuantidade_LostFocus()

    txtQuantidade.Text = Numeros(txtQuantidade.Text)

End Sub

Private Sub txtReferencia_GotFocus()

    txtReferencia.SelStart = 0
    txtReferencia.SelLength = Len(txtReferencia.Text)

End Sub

Private Sub txtReferencia_KeyDown(KeyCode As Integer, shift As Integer)

    Dim Linha As Long

    Linha = grdItens.Row
    
    clsItensNota.SelecionaLinha KeyCode
    
    If grdItens.Row <> Linha Then
        DesceDados
        PreencheCapa
    End If

End Sub

Private Sub txtReferencia_KeyPress(KeyAscii As Integer)
    
    VerTecla KeyAscii

End Sub
Private Sub txtSerie_GotFocus()
    
    txtSerie.SelStart = 0
    txtSerie.SelLength = Len(txtSerie.Text)

End Sub

Private Sub txtSerie_LostFocus()
    
    If GetAsyncKeyState(vbKeyTab) <> 0 Then
        If Not Val(txtNotaFiscal.Text) > 0 Then
             MsgBox "Número de Nota Fiscal inválido.", vbExclamation, "Atenção"
             txtNotaFiscal.SetFocus
             Exit Sub
        End If
        
        txtSerie.Text = Trimar(txtSerie.Text)
        If txtSerie.Text = "" Then
             MsgBox "Digite a série da Nota Fiscal.", vbExclamation, "Atenção"
             txtSerie.SetFocus
             Exit Sub
        End If
        
        If Len(txtSerie.Text) <> 2 Then
            MsgBox "A Série da Nota Fiscal deve conter 2 dígitos!", vbExclamation, "Atenção"
            txtSerie.SetFocus
            Exit Sub
        End If
        
        If Val(txtNomeFantasia.Text) = 0 Or Trim(txtCGC.Text) = "" Then
             MsgBox "Informe o fornecedor.", vbExclamation, "Atenção"
             txtCGC.SetFocus
             Exit Sub
        End If
        
        CarregaNotaFiscal
        
        If txtCodigoOperacao.Text <> "" Then
           Call ConverteCodigoOperacaoNovo
           txtCodigoOperacao.Text = WCodigoOperacaoNovo
        End If
        
    Else
        cmbLocal.Enabled = True
    End If
txtCGC.Enabled = False
End Sub

Private Sub CarregaNotaFiscal()

    txtSerie.Text = UCase(txtSerie.Text)
    
    clsNotaFiscal.UseWhere = False
    clsNotaFiscal.SQL = "Select [Campos], CC_Situacao " _
                     & "From CapaNFCompra where CC_NotaFiscal = " _
                     & Val(txtNotaFiscal.Text) & " and CC_Serie = '" _
                     & Trim(txtSerie.Text) & "' and CC_Fornecedor = " _
                     & Val(txtNomeFantasia.Text)
    
    Screen.MousePointer = 11
    
    Pesquisou = False
    
    clsNotaFiscal.Ler
    
    If Pesquisou And Not SoExclusao Then
        Screen.MousePointer = 11
        
        If NotaJaExiste Then
            If grdItens.Rows = 2 And grdItens.TextMatrix(1, 0) = "" Then
                txtCodigoOperacao.Enabled = True
                cmbNaturezaOperacao.Enabled = True
            Else
                txtCodigoOperacao.Enabled = False
                cmbNaturezaOperacao.Enabled = False
            End If
            
'            clsItensNota.SQL = "Select CF_CodigoOperacaoAux as CodigoOperacao, PI_NumeroPedido, PI_Referencia, " _
                             & "PR_Descricao, '' as CI_Quantidade, '' as CI_PrecoUnitario, " _
                             & "'' as CI_AliquotaIPI, '' as CI_PercentualDesconto, '' as Total, PI_SaldoPedido " _
                             & "From Produto, ItemPedido, CodigoOperacao, CapaPedido, CapaNFCompra " _
                             & "Where PR_Referencia=PI_Referencia and " _
                             & "PC_NumeroPedido=PI_NumeroPedido and " _
                             & "PC_CodigoOperacao=CF_CodigoOperacao and " _
                             & "CF_CodigoOperacaoAux=CC_CodigoOperacao and " _
                             & "PI_Situacao='A' and CC_NotaFiscal=" & Val(txtNotaFiscal.Text) _
                             & " and CC_Serie='" & txtSerie.Text & "' and " _
                             & "CC_Fornecedor=" & Val(txtNomeFantasia.Text) & " and " _
                             & "PI_Filial = '" & cmbLocal.Text & "' and " _
                             & "PC_CodigoFornecedor = " & Val(txtNomeFantasia.Text) _
                             & " order by PI_NumeroPedido, PR_Descricao"
            
            clsItensNota.SQL = "Select PC_LojaReserva,PR_Residencia,PC_CodigoOperacao as CodigoOperacao, PI_NumeroPedido, PI_Referencia, " _
                             & "PR_Descricao, '' as CI_Quantidade, '' as CI_PrecoUnitario, " _
                             & "PI_AliquotaIPI as CI_AliquotaIPI, PI_PercentualDesconto as CI_PercentualDesconto, '' as Total, PI_SaldoPedido, PC_NaturezaOperacao as NaturezaOperacao,CO_Nome as Comprador,PC_SemanaEntrega as Entrega,CP_Descricao as CondPag,PC_DespesaFrete as Frete " _
                             & "From CondicaoPagto,Produto, ItemPedido, CapaPedido, CapaNFCompra, Comprador, Fornecedor " _
                             & "Where CP_CodigoCondicao = PC_CondicaoPagamento and PR_Referencia=PI_Referencia and " _
                             & "CP_VendaCompra = 'C' and PC_NumeroPedido=PI_NumeroPedido and " _
                             & "PI_Situacao='A' and CC_NotaFiscal=" & Val(txtNotaFiscal.Text) _
                             & " and CC_Serie='" & txtSerie.Text & "' and " _
                             & "CC_Fornecedor=" & Val(txtNomeFantasia.Text) & " and " _
                             & "PI_Filial = '" & cmbLocal.Text & "' and PC_NumeroPedido = PI_NumeroPedido and PC_CodigoComprador = CO_CodigoComprador and " _
                             & "PC_FornecedorRecebimento = FO_FornecedorRecebimento and FO_CodigoFornecedor = " & Val(txtNomeFantasia.Text) _
                             & " order by PI_NumeroPedido, PR_Descricao"
        
        Else
            txtCodigoOperacao.Enabled = True
            cmbNaturezaOperacao.Enabled = True
            
            clsItensNota.SQL = "Select PC_LojaReserva,PR_Residencia,PC_CodigoOperacao as CodigoOperacao, PI_NumeroPedido, PI_Referencia, " _
                             & "PR_Descricao, '' as CI_Quantidade, '' as CI_PrecoUnitario, " _
                             & "PI_AliquotaIPI as CI_AliquotaIPI, PI_PercentualDesconto as CI_PercentualDesconto, '' as Total, PI_SaldoPedido, PC_NaturezaOperacao as NaturezaOperacao,CO_Nome as Comprador,PC_SemanaEntrega as Entrega,CP_Descricao as CondPag,PC_DespesaFrete as Frete " _
                             & " From CondicaoPagto,Produto, ItemPedido, CapaPedido, Comprador, Fornecedor " _
                             & "where PR_Referencia = PI_Referencia and " _
                             & "CP_VendaCompra = 'C' and PC_CondicaoPagamento = CP_CodigoCondicao and PC_NumeroPedido = PI_NumeroPedido and " _
                             & "PI_Situacao = 'A' and CO_CodigoComprador = PC_CodigoComprador and " _
                             & "PI_Filial = '" & cmbLocal.Text & "' and " _
                             & "PC_FornecedorRecebimento = FO_FornecedorRecebimento and FO_CodigoFornecedor = " & Val(txtNomeFantasia.Text) _
                             & " order by PI_NumeroPedido, PR_Descricao"
        End If
    
        LinhaCFO = -1
        clsItensNota.Preencher
        
        If grdItens.Rows > 2 And grdItens.TextMatrix(1, 0) = "" Then
            grdItens.RemoveItem 1
        End If
        
        RemoveRepetidos
    End If
    
    ObtemObservacoes
    
    If grdItens.TextMatrix(1, 0) <> "" Then
        txtCGC.Enabled = False
        txtNotaFiscal.Enabled = False
        txtSerie.Enabled = False
        cmbLocal.Enabled = False
        
        If NotaJaExiste Then
            MsgBox "Serão removidos, agora, os ítens que possuem natureza de operação diferente da selecionada.", vbInformation, "Informação"
            FiltraCodigoOperacao
            
            If SoExclusao = False Then
                grdItens.Redraw = True
                grdItens.Row = 1
                grdItens.Col = 0
                grdItens.ColSel = grdItens.Cols - 1
                grdItens.SetFocus
                DesceDados
                PreencheCapa
                txtQuantidade.SetFocus
            End If
        End If
    End If
    
    If NotaJaExiste And Not SoExclusao Then
        cmdVencimentos.Enabled = True
    Else
        cmdVencimentos.Enabled = False
    End If
    
    lblPedido.Visible = True
    lblComprador.Visible = True
    lblCFOP.Visible = True
    lblFrete.Visible = True
    lblEntrega.Visible = True
    lblEntrega2.Visible = True
    lblCondPagto.Visible = True
    lblCondPagto2.Visible = True
    
    Screen.MousePointer = 0

End Sub

Private Sub ObtemObservacoes()

    Dim rdoObs As New ADODB.Recordset
    Dim Apontador As Long
    Dim Maximo As Long
    Dim NumPed As Long
    Dim PedidoAnterior As Long
    Dim Pedidos As String
    
    Maximo = grdItens.Rows - 1
    
    On Error Resume Next
    
    ReDim Observacoes(1 To 1) As ObservacaoPedido
    
    If grdItens.TextMatrix(1, 0) <> "" Then
       ' Status "Obtendo observações para os Pedidos..."
        
        PedidoAnterior = 0
        Pedidos = ""
        For Apontador = 1 To Maximo Step 1
            NumPed = Val(grdItens.TextMatrix(Apontador, 1))
            If NumPed <> PedidoAnterior Then
                PedidoAnterior = NumPed
                Pedidos = Pedidos & "," & NumPed
            End If
        Next Apontador
        Pedidos = Mid(Pedidos, 2)
    
        If Pedidos <> "" Then
        
        SQL = "Select PC_NumeroPedido, PC_ObservacaoAlmox from CapaPedido Where PC_NumeroPedido in (" & Pedidos & ") order by PC_NumeroPedido"
         rdoObs.CursorLocation = adUseClient
         rdoObs.Open SQL, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
          '  Set rdoObs = rdoCnSupBatch.OpenResultset("Select PC_NumeroPedido, PC_ObservacaoAlmox from CapaPedido Where PC_NumeroPedido in (" & Pedidos & ") order by PC_NumeroPedido", Options:=rdExecDirect)
            Apontador = 1
            Do While Not rdoObs.EOF
                ReDim Preserve Observacoes(1 To Apontador) As ObservacaoPedido
                Observacoes(Apontador).NumeroPedido = rdoObs("PC_NumeroPedido")
                Observacoes(Apontador).Obs = rdoObs("PC_ObservacaoAlmox")
                Apontador = Apontador + 1
                rdoObs.MoveNext
            Loop
            rdoObs.Close
        End If
        
       ' Status "Pronto."
    End If
    

End Sub

Private Sub txtValorICMS_GotFocus()
    
    txtValorICMS.SelStart = 0
    txtValorICMS.SelLength = Len(txtValorICMS.Text)

End Sub
Private Sub txtValorICMS_KeyPress(KeyAscii As Integer)

    VerteclaVirgula Me.ActiveControl, KeyAscii

End Sub

Private Sub txtValorICMS_LostFocus()

    txtValorICMS.Text = Format(txtValorICMS.Text, "###,###,###,###0.00")

    If Val(ConverteVirgula(Format(txtaliquotaicms.Text, "0.00"))) = 0 And Val(ConverteVirgula(Format(txtValorICMS.Text, "0.00"))) <> 0 Then
       MsgBox "Informar a alíquota de ICMS", vbCritical, "Atenção"
       txtaliquotaicms.SetFocus
       Exit Sub
    End If
    
    If Val(ConverteVirgula(Format(txtaliquotaicms.Text, "0.00"))) <> 0 And Val(ConverteVirgula(Format(txtValorICMS.Text, "0.00"))) = 0 Then
       MsgBox "Informar o Valor do ICMS", vbCritical, "Atenção"
       txtValorICMS.SetFocus
       Exit Sub
    End If
    
End Sub

Private Sub txtValorICMSSubsTrib_GotFocus()
    
    txtValorICMSSubsTrib.SelStart = 0
    txtValorICMSSubsTrib.SelLength = Len(txtValorICMSSubsTrib.Text)

End Sub

Private Sub txtValorICMSSubsTrib_KeyPress(KeyAscii As Integer)

    VerteclaVirgula Me.ActiveControl, KeyAscii

End Sub

Private Sub txtValorICMSSubsTrib_LostFocus()

    txtValorICMSSubsTrib.Text = Format(txtValorICMSSubsTrib.Text, "###,###,###,###0.00")

End Sub

Private Sub txtValorIPI_GotFocus()
    
    txtValorIPI.SelStart = 0
    txtValorIPI.SelLength = Len(txtValorIPI.Text)

End Sub


Private Sub txtValorIPI_KeyPress(KeyAscii As Integer)

    VerteclaVirgula Me.ActiveControl, KeyAscii

End Sub

Private Sub txtValorIPI_LostFocus()

    txtValorIPI.Text = Format(txtValorIPI.Text, "###,###,###,###0.00")

End Sub

Private Sub txtValorMercadorias_GotFocus()
    
    txtValorMercadorias.SelStart = 0
    txtValorMercadorias.SelLength = Len(txtValorMercadorias.Text)

End Sub

Private Sub txtValorMercadorias_KeyPress(KeyAscii As Integer)

    VerteclaVirgula Me.ActiveControl, KeyAscii

End Sub

Private Sub txtValorMercadorias_LostFocus()

    txtValorMercadorias.Text = Format(txtValorMercadorias.Text, "###,###,###,###0.00")

End Sub

Private Sub txtValorTotalNota_Change()

    If Pesquisou Then
        CalculaBateNota
    End If

End Sub


Sub CalculaBateNota()

    Dim Total As String
        
    Total = Format(txtValorTotalNota.Text, "0.00")
    If Not IsNumeric(Total) Then
        BateNota = CDbl(Format(txtTotalCalculado.Text, "0.00"))
        txtBateNota.Text = Format(BateNota, "###,###,###,###0.00")
    Else
        BateNota = CDbl(Format(txtTotalCalculado.Text, "0.00")) - CDbl(Total)
        txtBateNota.Text = Format(BateNota, "###,###,###,###0.00")
    End If

End Sub

Private Sub txtValorTotalNota_GotFocus()

    txtValorTotalNota.SelStart = 0
    txtValorTotalNota.SelLength = Len(txtValorTotalNota)

End Sub

Private Sub txtValorTotalNota_KeyPress(KeyAscii As Integer)

    VerteclaVirgula Me.ActiveControl, KeyAscii

End Sub


Private Sub txtValorTotalNota_LostFocus()

    txtValorTotalNota.Text = Format(txtValorTotalNota.Text, "###,###,###,###0.00")

End Sub

Private Sub ConverteCodigoOperacaoNovo()

    Dim SQL As String
    Dim RdoCodOper As New ADODB.Recordset
    
    WCodigoOperacaoVelho = txtCodigoOperacao.Text
    
    SQL = "Select CN_CodigoOperacaoNovo from CodigoOperacaoNovo where CN_CodigoOperacaoAntigo=" & txtCodigoOperacao.Text & ""
    Set RdoCodOper = rdoCnSup.OpenResultset(SQL)
    
    If Not RdoCodOper.EOF Then
       WCodigoOperacaoNovo = RdoCodOper("CN_CODIGOOPERACAONOVO")
    Else
       WCodigoOperacaoNovo = txtCodigoOperacao.Text
    End If
   
    MontaComboNatureza Val(WCodigoOperacaoVelho), cmbNaturezaOperacao
    
    If cmbNaturezaOperacao.ListCount > 0 Then
        cmbNaturezaOperacao.ListIndex = 0
    End If

End Sub
'22/09/2006
'Por Celso
'Faz distribuição mesmo se for Distribuidor



