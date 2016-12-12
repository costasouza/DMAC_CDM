VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Begin VB.Form frmEncerraNFOutrasOperacoes 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Encerra Nota Fiscal Outras Operações"
   ClientHeight    =   7770
   ClientLeft      =   1140
   ClientTop       =   2160
   ClientWidth     =   16170
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   7770
   ScaleWidth      =   16170
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      Height          =   45
      Left            =   150
      ScaleHeight     =   45
      ScaleWidth      =   14910
      TabIndex        =   38
      Top             =   6810
      Width           =   14910
   End
   Begin VB.Frame fraRemetente 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   150
      TabIndex        =   27
      Top             =   1710
      Width           =   14910
      Begin VB.OptionButton optCliente 
         BackColor       =   &H00404040&
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Left            =   2265
         TabIndex        =   0
         Top             =   90
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   930
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
         Left            =   100
         TabIndex        =   1
         Top             =   585
         Width           =   1110
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
         Left            =   1290
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   585
         Width           =   5130
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
         Left            =   6495
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   585
         Width           =   1680
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
         Left            =   8250
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   585
         Width           =   1860
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
         Left            =   10185
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   585
         Width           =   4605
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
         Left            =   100
         Locked          =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   32
         Top             =   1230
         Width           =   4710
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
         Left            =   4890
         Locked          =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   31
         Top             =   1230
         Width           =   4425
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
         Left            =   9810
         Locked          =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   30
         Top             =   1230
         Width           =   1755
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
         Left            =   9390
         Locked          =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   29
         Top             =   1230
         Width           =   345
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
         Left            =   11640
         Locked          =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   28
         Top             =   1230
         Width           =   3150
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         Caption         =   "Código Cliente"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   100
         TabIndex        =   79
         Top             =   350
         Width           =   1110
      End
      Begin VB.Label lblFoneFaxDest 
         BackColor       =   &H00404040&
         Caption         =   "Fone/Fax"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   11640
         TabIndex        =   62
         Top             =   990
         Width           =   780
      End
      Begin VB.Label lblCNPJDest 
         BackColor       =   &H00404040&
         Caption         =   "CNPJ "
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   6495
         TabIndex        =   61
         Top             =   350
         Width           =   780
      End
      Begin VB.Label lblInscricaoEstadualDest 
         BackColor       =   &H00404040&
         Caption         =   "Inscrição Estadual"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   8250
         TabIndex        =   60
         Top             =   350
         Width           =   1590
      End
      Begin VB.Label lblEnderecoDest 
         BackColor       =   &H00404040&
         Caption         =   "Endereço "
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   10185
         TabIndex        =   59
         Top             =   350
         Width           =   780
      End
      Begin VB.Label lblMunicipioDest 
         BackColor       =   &H00404040&
         Caption         =   "Município"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   105
         TabIndex        =   58
         Top             =   990
         Width           =   780
      End
      Begin VB.Label lblBairroDest 
         BackColor       =   &H00404040&
         Caption         =   "Bairro"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   4890
         TabIndex        =   57
         Top             =   990
         Width           =   780
      End
      Begin VB.Label lblUFDest 
         BackColor       =   &H00404040&
         Caption         =   "UF"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   9390
         TabIndex        =   56
         Top             =   990
         Width           =   345
      End
      Begin VB.Label lblCepDest 
         BackColor       =   &H00404040&
         Caption         =   "CEP"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   9810
         TabIndex        =   55
         Top             =   990
         Width           =   780
      End
      Begin VB.Label lblDestinatario 
         BackColor       =   &H00404040&
         Caption         =   "Destinatário"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1290
         TabIndex        =   54
         Top             =   350
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
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   105
         TabIndex        =   53
         Top             =   105
         Width           =   2085
      End
   End
   Begin VB.Frame fraEmitente 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   150
      TabIndex        =   15
      Top             =   150
      Width           =   14910
      Begin VB.TextBox txtNroPedido 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00A3A3A3&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   12615
         Locked          =   -1  'True
         TabIndex        =   37
         Text            =   "999999"
         Top             =   345
         Width           =   1050
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
         Left            =   100
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   345
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
         Left            =   3900
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   345
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
         Left            =   5865
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   345
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
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1005
         Width           =   4800
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
         Left            =   12615
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1005
         Width           =   2160
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
         Left            =   10800
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1005
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
         Left            =   11220
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1005
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
         Left            =   4980
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1005
         Width           =   2970
      End
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
         Left            =   8025
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1005
         Width           =   2700
      End
      Begin VB.TextBox txtNotaFiscal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00A3A3A3&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         Left            =   13725
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "999999"
         Top             =   345
         Width           =   1050
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
         Left            =   7740
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   345
         Width           =   1095
      End
      Begin VB.Label lblInscricaoEstadual 
         BackColor       =   &H00404040&
         Caption         =   "Inscrição Estadual"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   5865
         TabIndex        =   52
         Top             =   100
         Width           =   1335
      End
      Begin VB.Label lblEmissao 
         BackColor       =   &H00404040&
         Caption         =   "Emissão"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   7740
         TabIndex        =   51
         Top             =   100
         Width           =   630
      End
      Begin VB.Label lblPedido 
         BackColor       =   &H00404040&
         Caption         =   "Pedido"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   12615
         TabIndex        =   50
         Top             =   105
         Width           =   780
      End
      Begin VB.Label lblNotaFiscal 
         BackColor       =   &H00404040&
         Caption         =   "Nota Fiscal"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   13725
         TabIndex        =   49
         Top             =   105
         Width           =   870
      End
      Begin VB.Label lblEndereco 
         BackColor       =   &H00404040&
         Caption         =   "Endereço"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   105
         TabIndex        =   48
         Top             =   765
         Width           =   1080
      End
      Begin VB.Label lblMunicipio 
         BackColor       =   &H00404040&
         Caption         =   "Município"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   4980
         TabIndex        =   47
         Top             =   765
         Width           =   1080
      End
      Begin VB.Label lblBairro 
         BackColor       =   &H00404040&
         Caption         =   "Bairro"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   8025
         TabIndex        =   46
         Top             =   765
         Width           =   1080
      End
      Begin VB.Label lblUf 
         BackColor       =   &H00404040&
         Caption         =   "UF"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   10800
         TabIndex        =   45
         Top             =   765
         Width           =   315
      End
      Begin VB.Label lblCep 
         BackColor       =   &H00404040&
         Caption         =   "CEP"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   11220
         TabIndex        =   44
         Top             =   765
         Width           =   435
      End
      Begin VB.Label lblFoneFax 
         BackColor       =   &H00404040&
         Caption         =   "Fone/Fax"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   12615
         TabIndex        =   43
         Top             =   765
         Width           =   1215
      End
      Begin VB.Label lblCnpj 
         BackColor       =   &H00404040&
         Caption         =   "CNPJ"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   3900
         TabIndex        =   42
         Top             =   100
         Width           =   1080
      End
      Begin VB.Label lblEmitente 
         BackColor       =   &H00404040&
         Caption         =   "Emitente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   100
         TabIndex        =   41
         Top             =   100
         Width           =   1530
      End
   End
   Begin VB.Frame fraTotalNF 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Totais da Nota Fiscal"
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
      Height          =   1020
      Left            =   150
      TabIndex        =   13
      Top             =   5670
      Width           =   14910
      Begin VB.Frame fraImpostos 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   930
         Left            =   6420
         TabIndex        =   71
         Top             =   0
         Width           =   8640
         Begin VB.TextBox txtTotalNFInf 
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
            Left            =   7035
            TabIndex        =   77
            Top             =   585
            Width           =   1350
         End
         Begin VB.TextBox txtValorIPI 
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
            Left            =   5625
            TabIndex        =   11
            Top             =   585
            Width           =   1350
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
            TabIndex        =   7
            Top             =   585
            Width           =   1320
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
            Left            =   1500
            TabIndex        =   8
            Top             =   585
            Width           =   1320
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
            Left            =   2880
            TabIndex        =   9
            Top             =   585
            Width           =   1275
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
            Left            =   4215
            TabIndex        =   10
            Top             =   585
            Width           =   1350
         End
         Begin VB.Label lblTotalNFInf 
            BackColor       =   &H00404040&
            Caption         =   "Total NF Informado"
            ForeColor       =   &H00E0E0E0&
            Height          =   195
            Left            =   7035
            TabIndex        =   78
            Top             =   350
            Width           =   1365
         End
         Begin VB.Label lblValorIPI 
            BackColor       =   &H00404040&
            Caption         =   "Valor IPI"
            ForeColor       =   &H00E0E0E0&
            Height          =   195
            Left            =   5625
            TabIndex        =   76
            Top             =   350
            Width           =   1155
         End
         Begin VB.Label lblValorICMSST 
            BackColor       =   &H00404040&
            Caption         =   "Valor ICMS ST"
            ForeColor       =   &H00E0E0E0&
            Height          =   195
            Left            =   4215
            TabIndex        =   75
            Top             =   350
            Width           =   1155
         End
         Begin VB.Label lblBaseCalcST 
            BackColor       =   &H00404040&
            Caption         =   "Base Cálculo ST"
            ForeColor       =   &H00E0E0E0&
            Height          =   195
            Left            =   2850
            TabIndex        =   74
            Top             =   350
            Width           =   1305
         End
         Begin VB.Label lblValorICMS 
            BackColor       =   &H00404040&
            Caption         =   "Valor ICMS"
            ForeColor       =   &H00E0E0E0&
            Height          =   195
            Left            =   1500
            TabIndex        =   73
            Top             =   350
            Width           =   885
         End
         Begin VB.Label lblBaseCalcICMS 
            BackColor       =   &H00404040&
            Caption         =   "Base Calc ICMS"
            ForeColor       =   &H00E0E0E0&
            Height          =   195
            Left            =   120
            TabIndex        =   72
            Top             =   350
            Width           =   1230
         End
      End
      Begin VB.CheckBox chkInformarImpostos 
         BackColor       =   &H00404040&
         Caption         =   "Informar Impostos"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   4785
         TabIndex        =   70
         Top             =   630
         Width           =   1920
      End
      Begin VB.CheckBox chkImpostosNao 
         BackColor       =   &H00404040&
         Caption         =   "Não Calcular Impostos"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2820
         TabIndex        =   69
         Top             =   630
         Width           =   1920
      End
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   6
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
         Left            =   100
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   585
         Width           =   1290
      End
      Begin VB.Label lblValorMercadoria 
         BackColor       =   &H00404040&
         Caption         =   "Valor Mercadoria"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   105
         TabIndex        =   68
         Top             =   350
         Width           =   1275
      End
      Begin VB.Label lblTotalNF 
         BackColor       =   &H00404040&
         Caption         =   "Total Nota Fiscal"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   1440
         TabIndex        =   67
         Top             =   350
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
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   100
         TabIndex        =   66
         Top             =   100
         Width           =   2040
      End
   End
   Begin VB.Frame FraCFOP 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   780
      Left            =   150
      TabIndex        =   12
      Top             =   3510
      Width           =   14910
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
         Left            =   2880
         TabIndex        =   3
         Text            =   "Combo2"
         Top             =   345
         Width           =   3930
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
         ItemData        =   "frmEncerraNFOutrasOperacoes.frx":0000
         Left            =   100
         List            =   "frmEncerraNFOutrasOperacoes.frx":0002
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   345
         Width           =   2700
      End
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
         Height          =   315
         Left            =   6885
         TabIndex        =   4
         Top             =   345
         Width           =   7920
      End
      Begin VB.Label lblCarimbo 
         BackColor       =   &H00404040&
         Caption         =   "Carimbo Nota Fiscal"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   6885
         TabIndex        =   65
         Top             =   100
         Width           =   780
      End
      Begin VB.Label lblCFOP 
         BackColor       =   &H00404040&
         Caption         =   "Código da Operação Fiscal"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   2880
         TabIndex        =   64
         Top             =   100
         Width           =   2430
      End
      Begin VB.Label lblTipoES 
         BackColor       =   &H00404040&
         Caption         =   "Tipo de Entrada/Saida"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   100
         TabIndex        =   63
         Top             =   100
         Width           =   2310
      End
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdItens 
      Height          =   1170
      Left            =   150
      TabIndex        =   14
      Top             =   4395
      Width           =   14910
      _cx             =   26300
      _cy             =   2064
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
      BackColor       =   14737632
      ForeColor       =   4210752
      BackColorFixed  =   0
      ForeColorFixed  =   16777215
      BackColorSel    =   3421236
      ForeColorSel    =   16777215
      BackColorBkg    =   12632256
      BackColorAlternate=   12632256
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmEncerraNFOutrasOperacoes.frx":0004
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
      BackColorFrozen =   -2147483633
      ForeColorFrozen =   4210752
      WallPaperAlignment=   9
   End
   Begin DmacADMLoja.chameleonButton cmdRetorna 
      Height          =   510
      Left            =   13680
      TabIndex        =   39
      Top             =   6960
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
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   5263440
      MPTR            =   1
      MICON           =   "frmEncerraNFOutrasOperacoes.frx":01D6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin DmacADMLoja.chameleonButton cmdEncerraNF 
      Height          =   510
      Left            =   12240
      TabIndex        =   40
      Top             =   6960
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Encerra"
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
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   5263440
      MPTR            =   1
      MICON           =   "frmEncerraNFOutrasOperacoes.frx":01F2
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

Dim wPesquisar As Boolean
Dim wLoja As String
Dim wDentroForaEstado As String
Dim LojaOrigem As String
Dim SQL As String
Dim wEnable As String
Dim wCor As String
Dim wIdx As Integer

Private Sub chkImpostosNao_Click()
    chkInformarImpostos.Value = 0
    fraImpostos.Visible = False
End Sub

Private Sub chkInformarImpostos_Click()
   chkImpostosNao.Value = 0
   If chkInformarImpostos.Value = 1 Then
     fraImpostos.Visible = True
     txtBaseICMS.Text = ""
     txtValorICMS.Text = ""
     txtBaseCalculoST.Text = ""
     txtValorICMSST.Text = ""
     txtTotalNFInf.Text = ""
     txtValorIPI.Text = ""
     wBaseCalculoST = 0
     wValorICMSST = 0
     wValorIPI = 0
   Else
     fraImpostos.Visible = False
   End If
   
End Sub

Private Sub cmbTipoES_Click()
 Call CarregaCFOP
End Sub

Private Sub cmdEncerraNF_Click()

 If chkInformarImpostos.Value = 1 Then
    If Trim(txtBaseICMS.Text) = "" Or Trim(txtValorICMS.Text) = "" Or Trim(txtBaseCalculoST.Text) = "" Or _
         Trim(txtValorICMSST.Text) = "" Or Trim(txtTotalNFInf.Text) = "" Or Trim(txtValorIPI.Text) = "" Then
         MsgBox "Informar todos os impostos"
         Exit Sub
    End If
 End If
 
 If wOpNFETotal = True Then
    NroNotaFiscal = ExtraiSeqNEControle
 Else
    NroNotaFiscal = ExtraiSeqNotaControle
 End If
 txtNotaFiscal.Text = NroNotaFiscal
 Call FinalizaNF
 AtualizaRetaguarda = True
 Unload Me
 Unload frmNotaFiscalOutrasOperacoes
 frmControleDmacAdmLoja.lblNomeTelas.Caption = ""

End Sub

Private Sub cmdRetorna_Click()
 frmControleDmacAdmLoja.lblNomeTelas.Caption = ""
 Unload Me
End Sub

Private Sub Form_Activate()
    PegaNumeroPedido
    txtNroPedido.Text = NroPedido
End Sub

Private Sub Form_Load()

Call AjustaTela(Me)

fraImpostos.Visible = False
txtEmitente.Text = fDataExt(Date)
'frmEncerraNFOutrasOperacoes.top = (Screen.Height - frmEncerraNFOutrasOperacoes.Height) / 2
'frmEncerraNFOutrasOperacoes.left = (Screen.Width - frmEncerraNFOutrasOperacoes.Width) / 2

    
'frmEncerraNFOutrasOperacoes.top = 3070
'frmEncerraNFOutrasOperacoes.left = 90
wDentroForaEstado = "D"
chkImpostosNao.Value = 0
chkInformarImpostos.Value = 0

chamou = "frmEncerraNFOutrasOperacoes"
 'PegaNumeroPedido
 'txtNroPedido.Text = NroPedido
txtNotaFiscal.Text = ""
Call CarregaRemetente
Call CarregaTipoES


End Sub

Private Sub Label9_Click()
End Sub



Private Sub lblTotalNota_Click()

End Sub

Private Sub lblNotaFiscal_Click()
100
End Sub

Private Sub optCliente_Click()
  txtCodigoDestinatario.Text = 0
 'lblCodigo.Visible = True
  txtCodigoDestinatario.Visible = True
  txtCodigoDestinatario.SetFocus
  ControlaDestinatario "C"
End Sub

Function ControlaDestinatario(ByVal TipoDestinatario As String)
  If TipoDestinatario = "I" Then
     wEnable = "False"
     wCor = vbWhite
     
  Else
     wEnable = "True"
     wCor = &HA3A3A3
  End If
  txtDestinatario.Locked = wEnable
  
  txtCNPJDestinatario.Locked = wEnable

  txtInscricaoDestinatario.Locked = wEnable
  
  txtEnderecoDestinatario.Locked = wEnable
  
  txtMunicipioDestinatario.Locked = wEnable
 
  txtBairroDestinatario.Locked = wEnable
  
  txtUFDestinatario.Locked = wEnable
  
  txtCepDestinatario.Locked = wEnable
  
  txtFoneFaxDestinatario.Locked = wEnable
  
  txtDestinatario.BackColor = wCor
     txtCNPJDestinatario.BackColor = wCor
     txtEnderecoDestinatario.BackColor = wCor
     txtInscricaoDestinatario.BackColor = wCor
      txtMunicipioDestinatario.BackColor = wCor
      txtBairroDestinatario.BackColor = wCor
  txtUFDestinatario.BackColor = wCor
  txtCepDestinatario.BackColor = wCor
  txtFoneFaxDestinatario.BackColor = wCor
  

End Function
Private Sub CarregaRemetente()
    SQL = "Select * From LOJA WHERE LO_LOJA ='" & GLB_Loja & "'"
    adoLoja.CursorLocation = adUseClient
    adoLoja.Open SQL, ADO_Cn_Dmac_Loja, adOpenForwardOnly, adLockPessimistic
    
    txtEmitente.Text = adoLoja("LO_razao")
    txtCNPJEmitente.Text = adoLoja("LO_CGC")
    txtInscricaoEmitente.Text = adoLoja("LO_InscricaoEstadual")
    txtDataEmissao.Text = Format(Date, "ddmmyyyy")
    txtEnderecoEmitente.Text = adoLoja("LO_Endereco")
    txtMunicipioEmitente.Text = adoLoja("LO_Municipio")
    txtBairroEmitente.Text = adoLoja("LO_Bairro")
    txtUFEmitente.Text = adoLoja("LO_UF")
    txtCepEmitente.Text = adoLoja("LO_Cep")
    txtFoneFaxEmitente.Text = adoLoja("LO_Telefone") & "/" & adoLoja("LO_Fax")
    
    adoLoja.Close
 
End Sub

Private Sub txtCodigoDestinatario_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    ' If optFornecedor.Value = True Then
    '    SQL = "Select * From Fornecedor WHERE FO_CodigoFornecedor =" & txtCodigoDestinatario.Text
    '    adoFornecedor.CursorLocation = adUseClient
    'adoFornecedor.Open SQL, ADO_Cn_Dmac, adOpenForwardOnly, adLockPessimistic
       
    '    If adoFornecedor.EOF Then
    '       MsgBox "Fornecedor não Cadastrado ", vbInformation, "Atenção"
    '       txtCodigoDestinatario.SelStart = 0
    '       txtCodigoDestinatario.SelLength = Len(txtCodigoDestinatario.Text)
    '       txtCodigoDestinatario.SetFocus
    '       adoFornecedor.Close
    '    Else
    '       txtDestinatario.Text = adoFornecedor("FO_RazaoSocial")
    '       txtCNPJDestinatario.Text = adoFornecedor("FO_CGC")
    '       txtInscricaoDestinatario.Text = adoFornecedor("FO_InscricaoEstadual")
    '       txtEnderecoDestinatario.Text = adoFornecedor("FO_Endereco")
    '       txtMunicipioDestinatario.Text = adoFornecedor("FO_Municipio")
    '       txtBairroDestinatario.Text = " "
    '       txtUFDestinatario.Text = adoFornecedor("FO_Estado")
    '       txtCepDestinatario.Text = adoFornecedor("FO_Cep")
    '       txtFoneFaxDestinatario.Text = adoFornecedor("FO_Telefone") & "/" & adoFornecedor("FO_Fax")
    '       adoFornecedor.Close
    '    End If
     If optCliente.Value = True Then
        SQL = "Select * From fin_Cliente WHERE CE_CodigoCliente =" & txtCodigoDestinatario.Text
         adoCliente.CursorLocation = adUseClient
        adoCliente.Open SQL, ADO_Cn_Dmac_Loja, adOpenForwardOnly, adLockPessimistic
       
        If adoCliente.EOF Then
           MsgBox "Cliente não Cadastrado ", vbInformation, "Atenção"
           txtCodigoDestinatario.SelStart = 0
           txtCodigoDestinatario.SelLength = Len(txtCodigoDestinatario.Text)
           txtCodigoDestinatario.SetFocus
           adoCliente.Close
        Else
           txtDestinatario.Text = adoCliente("CE_Razao")
           txtCNPJDestinatario.Text = adoCliente("CE_CGC")
           txtInscricaoDestinatario.Text = adoCliente("CE_InscricaoEstadual")
           txtEnderecoDestinatario.Text = adoCliente("CE_Endereco")
           txtMunicipioDestinatario.Text = adoCliente("CE_Municipio")
           txtBairroDestinatario.Text = adoCliente("CE_Bairro")
           txtUFDestinatario.Text = adoCliente("CE_Estado")
           txtCepDestinatario.Text = adoCliente("CE_Cep")
           txtFoneFaxDestinatario.Text = adoCliente("CE_Telefone") & "/" & adoCliente("CE_Fax")
            adoCliente.Close
        End If
     End If
     If txtUFDestinatario.Text <> txtUFEmitente.Text Then
        wDentroForaEstado = "F"
     Else
        wDentroForaEstado = "D"
     End If
     Call CarregaCFOP
     FraCFOP.Enabled = True
     cmdEncerraNF.Visible = True
  End If
End Sub
Private Sub CarregaTipoES()
    cmbTipoES.AddItem "100 - Entrada"
    cmbTipoES.AddItem "500 - Saída"

' SQL = "Select * from TipoEntradaSaidaCFOP "
' adotipo.CursorLocation = adUseClient
'    adotipo.Open SQL, ADO_Cn_Dmac, adOpenForwardOnly, adLockPessimistic

' If Not adotipo.EOF Then
'    Do While Not adotipo.EOF
'       cmbTipoES.AddItem adotipo("Tip_TipoCodigoES") & " - " & adotipo("Tip_Descricao")
'       adotipo.MoveNext
'    Loop
' End If
 cmbTipoES.ListIndex = 0
' adotipo.Close
End Sub

Private Sub CarregaCFOP()
 cmbCFOP.Clear
 SQL = "Select * from CFOPEntradaSaida Where cfo_entradasaida ='" & Mid(cmbTipoES.Text, 1, 3) & _
     "' and cfo_dentroforaestado = '" & wDentroForaEstado & "' order by cfo_descricaooperacao"
 adoCFOP.CursorLocation = adUseClient
 adoCFOP.Open SQL, ADO_Cn_Dmac_Loja, adOpenForwardOnly, adLockPessimistic

 If Not adoCFOP.EOF Then
    Do While Not adoCFOP.EOF
       cmbCFOP.AddItem adoCFOP("CFO_Codigo") & " - " & adoCFOP("CFO_DescricaoOperacao")
       adoCFOP.MoveNext
    Loop
    cmbCFOP.ListIndex = 0
 End If
 adoCFOP.Close
End Sub

Private Sub FinalizaNF()
 
    GetAsyncKeyState (vbKeyTab)
    
    wNroItens = 0
    If Mid(cmbTipoES.Text, 1, 3) = "502" Then
        wtiponota = "E"
    Else
        wtiponota = "S"
    End If
    
    CLIENTE = txtCodigoDestinatario.Text
  
    wTotNota = txtTotalNF.Text
    wVlrMercadoria = txtValormercadoria.Text
    wCFOP = Mid(Trim(cmbCFOP.Text), 1, 4)

   Do While grdItens.Rows - 1 > grdItens.Row
   
        grdItens.Row = grdItens.Row + 1
        wNroItens = wNroItens + 1
        wCodigoProduto = grdItens.TextMatrix(grdItens.Row, 0)
        wQtde = grdItens.TextMatrix(grdItens.Row, 2)
        wItemVenda = grdItens.TextMatrix(grdItens.Row, 3)
        wVlTotItem = wQtde * wItemVenda
   
   
        SQL = "Select PR_precovenda1,PR_icmPdv," _
             & "PR_Referencia,PR_descricao,* From Produtoloja ,ProdutoBarras " _
             & "Where PRB_Referencia = PR_Referencia and PRB_CodigoBarras='" & wCodigoProduto & "'"
            
                 
             adoProd.CursorLocation = adUseClient
             adoProd.Open SQL, ADO_Cn_Dmac_Loja, adOpenForwardOnly, adLockPessimistic
             
        If Not adoProd.EOF Then
            wICMS = Format(adoProd("PR_IcmsSaida"), "###,###,##0.00")
            wPLISTA = Format(adoProd("PR_PrecoVenda1"), "###,###,##0.00")
        End If
    
        adoProd.Close
        
        GravaItensPedido txtNroPedido.Text, "9"
      
   Loop
   
   CriaCapaPedido txtNroPedido.Text, "9"
   ADO_Cn_Dmac_Loja.BeginTrans
      SQL = "Insert into CarimboNotaFiscal(CNF_NumeroPed,CNF_Loja,CNF_serie,CNF_NF,CNF_Sequencia,CNF_Carimbo,CNF_TipoCarimbo)" _
        & "Values ( " & NroPedido & ",'" & GLB_Loja & "','" _
        & wSerie & "'," & NroNotaFiscal & "," & 1 & ",'" & txtCarimbo.Text & "','I')"
    
   ADO_Cn_Dmac_Loja.Execute (SQL)
   ADO_Cn_Dmac_Loja.CommitTrans
   
   
   
   EncerraVenda Val(txtNroPedido.Text), " ", 1


   If wOpNFETotal = False Then
        EmiteNotafiscal NroNotaFiscal, PegaSerieNota
   End If


  End Sub


