VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsflex7d.ocx"
Begin VB.Form frmEntradaMercadoria 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Consulta Entrada de Mercadoria"
   ClientHeight    =   7845
   ClientLeft      =   1260
   ClientTop       =   2010
   ClientWidth     =   15300
   LinkTopic       =   "Form16"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   15300
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "&H00C0C0C0&"
      Height          =   1020
      Left            =   10545
      TabIndex        =   21
      Top             =   150
      Width           =   4485
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   44
      Left            =   150
      ScaleHeight     =   45
      ScaleWidth      =   14880
      TabIndex        =   15
      Top             =   6810
      Width           =   14880
   End
   Begin VB.Frame fraQtdeItens 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   12255
      TabIndex        =   10
      Top             =   6075
      Width           =   2775
      Begin VB.TextBox txtQtdeItens 
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
         Left            =   1470
         TabIndex        =   11
         Top             =   165
         Width           =   1170
      End
      Begin VB.Label lblQtdeItens 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Quantidade Itens"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   150
         TabIndex        =   12
         Top             =   255
         Width           =   1215
      End
   End
   Begin VB.Frame fraPesquisa 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Pesquisa"
      ForeColor       =   &H00FF0000&
      Height          =   1020
      Left            =   150
      TabIndex        =   9
      Top             =   150
      Width           =   10410
      Begin VB.OptionButton optloja 
         BackColor       =   &H00404040&
         Caption         =   "Loja"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   9570
         TabIndex        =   7
         Top             =   150
         Width           =   675
      End
      Begin VB.OptionButton optReferencia 
         BackColor       =   &H00404040&
         Caption         =   "Referência"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   8400
         TabIndex        =   6
         Top             =   150
         Width           =   1095
      End
      Begin MSMask.MaskEdBox mskPeriodoFim 
         Height          =   315
         Left            =   2445
         TabIndex        =   1
         Top             =   540
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   10724259
         ForeColor       =   4210752
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskPeriodoIni 
         Height          =   315
         Left            =   825
         TabIndex        =   0
         Top             =   540
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   10724259
         ForeColor       =   4210752
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.OptionButton optcomprador 
         BackColor       =   &H00404040&
         Caption         =   "Comprador"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   7260
         TabIndex        =   5
         Top             =   150
         Width           =   1155
      End
      Begin VB.OptionButton optfornecedor 
         BackColor       =   &H00404040&
         Caption         =   "Fornecedor"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   6015
         TabIndex        =   4
         Top             =   150
         Width           =   1110
      End
      Begin VB.OptionButton optlinha 
         BackColor       =   &H00404040&
         Caption         =   "Linha"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   5175
         TabIndex        =   3
         Top             =   150
         Width           =   795
      End
      Begin VB.OptionButton opttodos 
         BackColor       =   &H00404040&
         Caption         =   "Todos"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   4335
         TabIndex        =   2
         Top             =   150
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.TextBox txtPesquisa 
         BackColor       =   &H00A3A3A3&
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   4335
         MaxLength       =   7
         TabIndex        =   8
         Top             =   540
         Width           =   5850
      End
      Begin VB.Label lblPesquisa 
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
         Left            =   150
         TabIndex        =   20
         Top             =   150
         Width           =   855
      End
      Begin VB.Label lblPeriodoFim 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "à"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   2235
         TabIndex        =   14
         Top             =   600
         Width           =   90
      End
      Begin VB.Label lblPeriodoIni 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Período"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   150
         TabIndex        =   13
         Top             =   600
         Width           =   570
      End
   End
   Begin CentroDeDistribuicao.chameleonButton cmdPesquisa 
      Height          =   510
      Left            =   12180
      TabIndex        =   16
      Top             =   6945
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
      MICON           =   "EntrMerc.frx":0000
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
      TabIndex        =   17
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
      MICON           =   "EntrMerc.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdNovaConsulta 
      Height          =   510
      Left            =   10740
      TabIndex        =   18
      Top             =   6945
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Nova Consulta"
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
      MICON           =   "EntrMerc.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdItemNota 
      Height          =   4605
      Left            =   150
      TabIndex        =   19
      Top             =   1335
      Width           =   14880
      _cx             =   26247
      _cy             =   8123
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
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   2
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"EntrMerc.frx":0054
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
End
Attribute VB_Name = "frmEntradaMercadoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents clsEntraMerc As ControlaGrid
Attribute clsEntraMerc.VB_VarHelpID = -1

Private Sub clsEntraMerc_RegistroNaoExiste()
   Screen.MousePointer = 0
   MsgBox "Nenhum item foi encontrado.", vbCritical, "Atenção"
   Limpar
End Sub

Private Sub cmdNovaConsulta_Click()
   Limpar
End Sub

Sub Limpar()
   mskPeriodoIni.Text = "__/__/____"
   mskPeriodoFim.Text = "__/__/____"
   optTodos.Value = True
   txtPesquisa.Text = ""
   txtQtdeItens.Text = ""
   clsEntraMerc.Clear
   mskPeriodoIni.SetFocus
   cmdPesquisa.Enabled = False
End Sub

Private Sub cmdPesquisa_Click()
   Dim Entrada As rdoResultset
   Dim wItens As Integer
         
   Screen.MousePointer = 11
    
   If mskPeriodoIni.Text <> "__/__/____" Then
      If mskPeriodoFim.Text <> "__/__/____" Then
         If (txtPesquisa.Text <> "" And optTodos.Value = False) Or (optTodos.Value = True) Then
            ClausulaWhere = " and CI_DataEntrada between '" & Format((mskPeriodoIni.Text), "mm/dd/yyyy") & "' and '" & Format((mskPeriodoFim.Text), "mm/dd/yyyy") & "' "
            
            If optLinha.Value = True Then
               If Len(Trim(txtPesquisa.Text)) <> 0 Then
                  ClausulaWhere = ClausulaWhere & " and PR_Linha = " & Trim(txtPesquisa.Text) & ""
               End If
            ElseIf optFornecedor.Value = True Then
               If Len(txtPesquisa.Text) <> 0 Then
                  ClausulaWhere = ClausulaWhere & " and CI_Fornecedor = " & Val(txtPesquisa.Text) & ""
               End If
            ElseIf optcomprador.Value = True Then
               If Len(txtPesquisa.Text) <> 0 Then
                  ClausulaWhere = ClausulaWhere & " and PR_Comprador = " & Trim(txtPesquisa.Text) & ""
               End If
            ElseIf optReferencia.Value = True Then
               If Len(txtPesquisa.Text) <> 0 Then
                  ClausulaWhere = ClausulaWhere & " and CI_Referencia ='" & Trim(txtPesquisa.Text) & "'"
               End If
            ElseIf optloja.Value = True Then
               If Len(txtPesquisa.Text) <> 0 Then
                  ClausulaWhere = ClausulaWhere & " and CC_Loja = '" & Trim(txtPesquisa.Text) & "'"
               End If
            End If
             
            clsEntraMerc.sql = "Select CI_Referencia, CI_Quantidade, CC_NotaFiscal, CC_Serie, CI_DataEntrada, " _
                             & "CI_DataLiberacao, CI_DataTransferencia, CI_PrecoUnitario, CI_NovoCusto, ci_fornecedor," _
                             & "CC_Loja, CC_CodigoOperacao, PR_Comprador, PR_Descricao, PR_Classe from Produto, " _
                             & "CapaNFCompra, ItemNFCompra where CC_NotaFiscal = CI_NotaFiscal and CC_Serie = CI_Serie and " _
                             & "CC_Fornecedor = CI_Fornecedor and PR_Referencia = CI_Referencia  " & ClausulaWhere _
                             & " order by CI_DataEntrada,ci_fornecedor"
            clsEntraMerc.Clear
            clsEntraMerc.Preencher
            
            grdItemNota.MergeCol(0) = True
            
            If grdItemNota.Rows > 2 And grdItemNota.TextMatrix(1, 0) = "" Then
               grdItemNota.RemoveItem 1
            End If
            
            wItens = grdItemNota.Rows - 1
            cmdPesquisa.Enabled = False
            Screen.MousePointer = 0
            txtQtdeItens = wItens
               
         Else
            MsgBox "Favor preencher dado à ser pesquisado", vbInformation, "Informação"
            Screen.MousePointer = 0
            txtPesquisa.SetFocus
            Exit Sub
         End If
      Else
         MsgBox "Favor preencher o campo período final", vbInformation, "Informação"
         Screen.MousePointer = 0
         mskPeriodoFim.SelStart = 0
         mskPeriodoFim.SelLength = Len(mskPeriodoFim.Text)
         mskPeriodoFim.SetFocus
         Exit Sub
      End If
   Else
      MsgBox "Favor preencher o campo período inicial", vbInformation, "Informação"
      Screen.MousePointer = 0
      mskPeriodoIni.SelStart = 0
      mskPeriodoIni.SelLength = Len(mskPeriodoIni.Text)
      mskPeriodoIni.SetFocus
      Exit Sub
   End If
End Sub

Private Sub cmdRetorna_Click()
    frmControleCD.lblNomeTelas.Caption = ""
   Unload Me
End Sub


Private Sub Form_Activate()
'JanelaTOP Me
carregarPosicaoTamanhoTela Me
   
End Sub


Private Sub Form_Load()
   

   Set clsEntraMerc = New ControlaGrid
   Set clsEntraMerc.ConexaoGrid = ADO_Cn_CD
   
   clsEntraMerc.NomeFormulario = Me.Name
   clsEntraMerc.NomeGrid = "grdItemNota"
   clsEntraMerc.Colunas = 15
   clsEntraMerc.LinhasVisiveis = 12
   
   clsEntraMerc.Cabecalho = "Referência; Código Barras; Quant.; Nota; Série; Fornecedor; Data Ent.; Data Lib.; Data Tran.; " _
                          & "Preço Unitário; Custo Unitário; Loja; CFO; Comprador; Descrição; Classe"
                          
   clsEntraMerc.Campos = "CI_Referencia; Prb_CodigoBarras; CI_Quantidade; CC_NotaFiscal; CC_Serie; CI_fornecedor; CI_DataEntrada; " _
                       & "CI_DataLiberacao; CI_DataTransferencia; CI_PrecoUnitario; CI_NovoCusto; " _
                       & "CC_Loja; CC_CodigoOperacao; PR_Comprador; PR_Descricao; PR_Classe"
                       
   clsEntraMerc.Formato = "Caractere; Caractere; Numero; Numero; Caractere; Numero; Data; Data; Data; " _
                          & "Decimal; Decimal; Caractere; Numero; Numero; Caractere; Caractere"
                          
   clsEntraMerc.Alinhamento = "Esquerda; Esquerda; Direita; Direita; Esquerda; Direita; Esquerda; Esquerda; Esquerda; " _
                          & "Direita; Direita; Esquerda; Direita; Direita; Esquerda; Esquerda"
                          
   clsEntraMerc.Tamanho = "1080; 1700; 700; 900; 540; 1300; 1200; 1200; 1200; 1350; 1350; 700; 540; 1000; 3800; 150"
   
   clsEntraMerc.MontaCabecalho
   mskPeriodoFim.Text = "__/__/____"
   mskPeriodoIni.Text = "__/__/____"
   cmdPesquisa.Enabled = False
   
End Sub

Private Sub mskPeriodoFim_LostFocus()
   If GetAsyncKeyState(vbKeyTab) <> 0 Then
      If mskPeriodoFim.Text <> "__/__/____" Then
         If DateDiff("d", Format(mskPeriodoFim.Text, "dd/mm/yyyy"), Format(mskPeriodoIni.Text, "dd/mm/yyyy")) > 0 Then
            MsgBox "Data final não pode ser anterior à data inicial do período", vbInformation, "Informação"
            mskPeriodoFim.SelStart = 0
            mskPeriodoFim.SelLength = Len(mskPeriodoFim.Text)
            mskPeriodoFim.SetFocus
            Exit Sub
         End If
      Else
         MsgBox "Favor preencher o campo período final", vbInformation, "Informação"
         Screen.MousePointer = 0
         mskPeriodoFim.SelStart = 0
         mskPeriodoFim.SelLength = Len(mskPeriodoFim.Text)
         mskPeriodoFim.SetFocus
         Exit Sub
      End If
   End If
End Sub

Private Sub mskperiodoini_lostfocus()
   If GetAsyncKeyState(vbKeyTab) <> 0 Then
      If mskPeriodoIni.Text = "__/__/____" Then
         MsgBox "Favor preencher período inicial", vbInformation, "Informação"
         mskPeriodoIni.SetFocus
         Exit Sub
      End If
   End If
End Sub

Private Sub optComprador_LostFocus()
If GetAsyncKeyState(vbKeyTab) <> 0 Then
   txtPesquisa.MaxLength = 2
   txtPesquisa.SetFocus
End If
End Sub

Private Sub optFornecedor_LostFocus()
If GetAsyncKeyState(vbKeyTab) <> 0 Then
   txtPesquisa.MaxLength = 4
   txtPesquisa.SetFocus
End If
End Sub

Private Sub optLinha_LostFocus()
If GetAsyncKeyState(vbKeyTab) <> 0 Then
   txtPesquisa.MaxLength = 2
   txtPesquisa.SetFocus
End If
End Sub

Private Sub optLoja_LostFocus()
If GetAsyncKeyState(vbKeyTab) <> 0 Then
   txtPesquisa.MaxLength = 5
   txtPesquisa.SetFocus
End If
End Sub

Private Sub optReferencia_LostFocus()
If GetAsyncKeyState(vbKeyTab) <> 0 Then
   txtPesquisa.MaxLength = 7
   txtPesquisa.SetFocus
End If
End Sub
Private Sub opttodos_lostfocus()
If GetAsyncKeyState(vbKeyTab) <> 0 Then
   txtPesquisa.MaxLength = 0
   cmdPesquisa.Enabled = True
   cmdPesquisa.SetFocus
End If
End Sub

Private Sub txtPesquisa_LostFocus()
   If GetAsyncKeyState(vbKeyTab) <> 0 Then
      cmdPesquisa.Enabled = True
      cmdPesquisa.SetFocus
   End If
End Sub

Private Sub grditemnota_dblclick()

   ''frmConsCapaNf.Show 1
   
   
End Sub
