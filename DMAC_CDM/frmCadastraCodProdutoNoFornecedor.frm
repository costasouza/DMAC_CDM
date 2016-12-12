VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Begin VB.Form frmCadastraCodProdutoNoFornecedor 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Cadastro de Produto no Fornecedor"
   ClientHeight    =   7980
   ClientLeft      =   3615
   ClientTop       =   2640
   ClientWidth     =   15810
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   15810
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   44
      Left            =   150
      ScaleHeight     =   45
      ScaleWidth      =   14880
      TabIndex        =   8
      Top             =   6810
      Width           =   14880
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   150
      TabIndex        =   1
      Top             =   5865
      Width           =   14900
      Begin VB.TextBox txtCodigoFornecedor 
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
         Height          =   285
         Left            =   5820
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   390
         Width           =   1890
      End
      Begin VB.TextBox txtDescricaoFornecedor 
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
         Height          =   285
         Left            =   1515
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   390
         Width           =   4230
      End
      Begin VB.TextBox txtReferencia 
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
         Height          =   285
         Left            =   150
         TabIndex        =   3
         Top             =   390
         Width           =   1290
      End
      Begin VB.Label lblCodigoFornecedor 
         BackColor       =   &H00404040&
         Caption         =   "Código do Fornecedor"
         ForeColor       =   &H00C0C0C0&
         Height          =   210
         Left            =   5820
         TabIndex        =   7
         Top             =   150
         Width           =   1890
      End
      Begin VB.Label lblDescricaoFornecedor 
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição do Fornecedor"
         ForeColor       =   &H00C0C0C0&
         Height          =   210
         Left            =   1515
         TabIndex        =   6
         Top             =   150
         Width           =   2235
      End
      Begin VB.Label lblReferencia 
         BackColor       =   &H00404040&
         Caption         =   "Referência"
         ForeColor       =   &H00C0C0C0&
         Height          =   200
         Left            =   150
         TabIndex        =   2
         Top             =   150
         Width           =   1005
      End
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdProduto 
      Height          =   5595
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   14895
      _cx             =   26273
      _cy             =   9869
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmCadastraCodProdutoNoFornecedor.frx":0000
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
   Begin CentroDeDistribuicao.chameleonButton cmdGrava 
      Height          =   510
      Left            =   12180
      TabIndex        =   9
      Top             =   6945
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Grava"
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
      MICON           =   "frmCadastraCodProdutoNoFornecedor.frx":00B0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdEncerra 
      Height          =   510
      Left            =   13620
      TabIndex        =   10
      Top             =   6945
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
      FCOL            =   16423203
      FCOLO           =   16423203
      MCOL            =   5263440
      MPTR            =   1
      MICON           =   "frmCadastraCodProdutoNoFornecedor.frx":00CC
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
Attribute VB_Name = "frmCadastraCodProdutoNoFornecedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdEncerra_Click()
    telaChamou = "frmCadastraCodProdutoNoFornecedor"

    Unload Me
    
End Sub

Private Sub cmdGrava_Click()
    If txtDescricaoFornecedor.Text = "" Or txtCodigoFornecedor.Text = "" Then
        MsgBox "Não há item selecionado!", vbCritical, "ATENÇÃO"
        Exit Sub
    End If
    If txtReferencia.Text = "" Then
        MsgBox "Preencha o campo referência!", vbCritical, "ATENÇÃO"
        txtReferencia.SetFocus
        Exit Sub
    End If
    
    sql = "Select * from produto where pr_referencia = '" & txtReferencia.Text & "'"
        adoXML.CursorLocation = adUseClient
        adoXML.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If adoXML.EOF Then
        MsgBox "Referência não cadastrada!", vbCritical, "ATENÇÃO"
        txtReferencia.SetFocus
    Else
        ADO_Cn_CDLocal.BeginTrans
            sql = "Update produto set pr_codigoprodutonofornecedor = '" & txtCodigoFornecedor & _
              "' where pr_referencia = '" & txtReferencia.Text & "'"
        ADO_Cn_CDLocal.Execute (sql)
        ADO_Cn_CDLocal.CommitTrans
            
        MsgBox "Cadastro realizado com sucesso!", vbInformation, "Cadastro Produto no Fornecedor"
        txtReferencia.Text = ""
        txtDescricaoFornecedor.Text = ""
        txtCodigoFornecedor.Text = ""
    End If
    
    adoXML.Close
End Sub


Private Sub Form_Load()
   JanelaTOP Me
   telaChamou = ""
   carregarPosicaoTamanhoTela Me
End Sub

Private Sub grdProduto_EnterCell()
    txtDescricaoFornecedor.Text = grdProduto.TextMatrix(grdProduto.row, 0)
    txtCodigoFornecedor.Text = grdProduto.TextMatrix(grdProduto.row, 1)
    txtReferencia.SetFocus
End Sub
