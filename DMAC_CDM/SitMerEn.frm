VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Begin VB.Form frmSitMercEnt 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Situação da Mercadoria de Entrada"
   ClientHeight    =   7470
   ClientLeft      =   -885
   ClientTop       =   2595
   ClientWidth     =   15195
   LinkTopic       =   "Form14"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   15195
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "&H00C0C0C0&"
      Height          =   6525
      Left            =   11595
      TabIndex        =   18
      Top             =   150
      Width           =   3450
   End
   Begin VB.Frame fraQtdeItens 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   8760
      TabIndex        =   15
      Top             =   6090
      Width           =   2685
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
         Left            =   1395
         TabIndex        =   16
         Top             =   135
         Width           =   1170
      End
      Begin VB.Label lblQtdeItens 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Quantidade Itens"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   75
         TabIndex        =   17
         Top             =   225
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00808080&
      Height          =   44
      Left            =   150
      ScaleHeight     =   45
      ScaleWidth      =   14880
      TabIndex        =   11
      Top             =   6810
      Width           =   14880
   End
   Begin VB.Frame fraOpcoes 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   915
      Left            =   150
      TabIndex        =   9
      Top             =   840
      Width           =   11295
      Begin VB.CheckBox CheckOrdenadaReferencia 
         BackColor       =   &H00404040&
         Caption         =   "Ordenada por Referência"
         Enabled         =   0   'False
         ForeColor       =   &H00C0C0C0&
         Height          =   210
         Left            =   7155
         TabIndex        =   7
         Top             =   600
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.CheckBox CheckOrdenaDescricao 
         BackColor       =   &H00404040&
         Caption         =   "Ordenada por Descrição"
         Enabled         =   0   'False
         ForeColor       =   &H00C0C0C0&
         Height          =   210
         Left            =   4935
         TabIndex        =   6
         Top             =   600
         Visible         =   0   'False
         Width           =   2040
      End
      Begin VB.OptionButton optPesquisa2 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Fornecedor"
         ForeColor       =   &H00C0C0C0&
         Height          =   210
         Index           =   2
         Left            =   3315
         TabIndex        =   4
         Top             =   195
         Width           =   1110
      End
      Begin VB.OptionButton optPesquisa2 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Referência"
         ForeColor       =   &H00C0C0C0&
         Height          =   210
         Index           =   1
         Left            =   1710
         TabIndex        =   3
         Top             =   195
         Width           =   1080
      End
      Begin VB.OptionButton optPesquisa2 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Comprador"
         ForeColor       =   &H00C0C0C0&
         Height          =   210
         Index           =   0
         Left            =   150
         TabIndex        =   2
         Top             =   195
         Width           =   1080
      End
      Begin VB.TextBox txtPesquisa 
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
         MaxLength       =   2
         TabIndex        =   5
         Top             =   150
         Width           =   4395
      End
   End
   Begin VB.Frame fraPesquisa 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Pesquisa"
      ForeColor       =   &H00FF0000&
      Height          =   525
      Left            =   150
      TabIndex        =   8
      Top             =   150
      Width           =   11295
      Begin VB.OptionButton optPesquisa1 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "A Transferir"
         ForeColor       =   &H00C0C0C0&
         Height          =   210
         Index           =   1
         Left            =   2925
         TabIndex        =   1
         Top             =   150
         Width           =   1110
      End
      Begin VB.OptionButton optPesquisa1 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "A Ser Aprovado Por Compras"
         ForeColor       =   &H00C0C0C0&
         Height          =   210
         Index           =   0
         Left            =   150
         TabIndex        =   0
         Top             =   150
         Value           =   -1  'True
         Width           =   2370
      End
   End
   Begin CentroDeDistribuicao.chameleonButton cmdPesquisa 
      Height          =   510
      Left            =   12180
      TabIndex        =   12
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
      MICON           =   "SitMerEn.frx":0000
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
      TabIndex        =   13
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
      MICON           =   "SitMerEn.frx":001C
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
      TabIndex        =   14
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
      MICON           =   "SitMerEn.frx":0038
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
      Height          =   3990
      Left            =   150
      TabIndex        =   10
      Top             =   1950
      Width           =   11295
      _cx             =   19923
      _cy             =   7038
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
      FixedCols       =   2
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"SitMerEn.frx":0054
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
Attribute VB_Name = "frmSitMercEnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents clsSitMerc As ControlaGrid
Attribute clsSitMerc.VB_VarHelpID = -1



Private Sub CheckOrdenadaReferencia_Click()
If CheckOrdenadaReferencia.Value = 1 Then
   CheckOrdenaDescricao.Enabled = False
ElseIf CheckOrdenadaReferencia.Value = 0 Then
   CheckOrdenaDescricao.Enabled = True
End If

End Sub

Private Sub CheckOrdenadaReferencia_LostFocus()
If CheckOrdenadaReferencia.Value = True Then
    cmdPesquisa.SetFocus
    CheckOrdenaDescricao.Value = False
    CheckOrdenaDescricao = 0
End If

End Sub


Private Sub CheckOrdenaDescricao_Click()
If CheckOrdenaDescricao.Value = 1 Then
   CheckOrdenadaReferencia.Enabled = False
ElseIf CheckOrdenaDescricao.Value = 0 Then
   CheckOrdenadaReferencia.Enabled = True
End If
End Sub

Private Sub CheckOrdenaDescricao_LostFocus()
If CheckOrdenaDescricao.Value = True Then
    cmdPesquisa.SetFocus
End If

End Sub

Private Sub clsSitMerc_Acabou()
   If grdItemNota.Rows > 2 And grdItemNota.TextMatrix(1, 0) = "" Then
      grdItemNota.RemoveItem 1
      txtQtdeItens.Text = grdItemNota.Rows - 1
   End If
End Sub

Private Sub clsSitMerc_RegistroNaoExiste()
   MsgBox "Nenhum registro encontrado", vbInformation, "Informação"
   cmdNovaConsulta_Click
End Sub

Private Sub cmdNovaConsulta_Click()
   fraOpcoes.Enabled = True
   txtPesquisa.Width = 4680
   optPesquisa1(0).Value = True
   optPesquisa2(0).Value = True
   txtPesquisa.Text = ""
   clsSitMerc.Clear
   txtQtdeItens.Text = ""
   optPesquisa1(0).SetFocus
   CheckOrdenadaReferencia.Value = 0
   CheckOrdenaDescricao.Value = 0
End Sub

Private Sub cmdPesquisa_Click()
   fraOpcoes.Enabled = False
   
   DefineOption
   
   If txtPesquisa.Text = "" Then
      MsgBox "Favor preencher o campo pesquisa", vbInformation, "Informação"
      fraOpcoes.Enabled = True
      txtPesquisa.SetFocus
      Exit Sub
   End If
    
   Screen.MousePointer = 11
   
   If optPesquisa1(0).Value = True Then
      If CheckOrdenaDescricao.Value = True Then
         clsSitMerc.sql = "Select CI_Referencia, CI_Quantidade, CI_NotaFiscal, CI_Serie, " _
                        & "CI_DataEntrada, CI_Situacao, PR_Descricao from ItemNFCompra, capanfcompra, " _
                        & "Produto where PR_Referencia=CI_Referencia and CI_Situacao='E' and ci_notafiscal=cc_notafiscal and ci_fornecedor=cc_fornecedor and cc_loja in ('ALM01','CD') and " _
                        & "" & ClausulaWhere & " order by PR_Descricao"
      Else
         clsSitMerc.sql = "Select CI_Referencia, CI_Quantidade, CI_NotaFiscal, CI_Serie, " _
                        & "CI_DataEntrada, CI_Situacao, PR_Descricao from ItemNFCompra,capanfcompra,  " _
                        & "Produto where PR_Referencia=CI_Referencia and CI_Situacao='E' and ci_notafiscal=cc_notafiscal and ci_fornecedor=cc_fornecedor and cc_loja in ('ALM01','CD') and " _
                        & "" & ClausulaWhere & " order by CI_Referencia"
      End If
      
      clsSitMerc.Clear
      clsSitMerc.Preencher
      
      grdItemNota.MergeCol(0) = True
      grdItemNota.MergeCol(1) = True
      
   Else
      If CheckOrdenaDescricao.Value = True Then
         clsSitMerc.sql = "Select CI_Referencia, CI_Quantidade, CI_NotaFiscal, CI_Serie, " _
                        & "CI_DataEntrada, CI_Situacao, PR_Descricao from ItemNFCompra,capanfcompra,  " _
                        & "Produto where PR_Referencia=CI_Referencia and CI_Situacao='L' and ci_notafiscal=cc_notafiscal and ci_fornecedor=cc_fornecedor and cc_loja in ('ALM01','CD') and " _
                        & "" & ClausulaWhere & " order by PR_Descricao"
      Else
         clsSitMerc.sql = "Select CI_Referencia, CI_Quantidade, CI_NotaFiscal, CI_Serie, " _
                        & "CI_DataEntrada, CI_Situacao, PR_Descricao from ItemNFCompra,capanfcompra,  " _
                        & "Produto where PR_Referencia=CI_Referencia and CI_Situacao='L' and ci_notafiscal=cc_notafiscal and ci_fornecedor=cc_fornecedor and cc_loja in ('ALM01','CD') and " _
                        & "" & ClausulaWhere & " order by CI_Referencia"
      End If
      
      clsSitMerc.Clear
      clsSitMerc.Preencher
      
      grdItemNota.MergeCol(0) = True
      grdItemNota.MergeCol(1) = True
      
   End If
   
   
   Screen.MousePointer = 0
   
End Sub

Sub DefineOption()
   ClausulaWhere = ""
  
   If optPesquisa2(0).Value = True Then
      If Len(Trim(txtPesquisa.Text)) > 0 Then
         If IsNumeric(Trim(txtPesquisa.Text)) Then
            ClausulaWhere = "PR_Comprador = " & Trim(txtPesquisa.Text) & " AND "
         End If
      End If
   End If
   
   If optPesquisa2(1).Value = True Then
      If Len(Trim(txtPesquisa.Text)) > 0 Then
         ClausulaWhere = ClausulaWhere & "PR_REFERENCIA = '" & Trim(txtPesquisa.Text) & "' AND "
      End If
   End If
   
   If optPesquisa2(2).Value = True Then
      If Len(Trim(txtPesquisa.Text)) > 0 Then
         ClausulaWhere = ClausulaWhere & "CI_Fornecedor = " & Trim(txtPesquisa.Text) & " AND "
      End If
   End If
   
   If txtPesquisa.Text <> "" Then
      ClausulaWhere = left(ClausulaWhere, Len(ClausulaWhere) - 4)
   End If

End Sub

Private Sub cmdRetorna_Click()
   frmControleCD.lblNomeTelas.Caption = ""
   Unload Me
End Sub

Private Sub Form_Activate()
carregarPosicaoTamanhoTela Me
'JanelaTOP Me
End Sub


Private Sub Form_Load()
   
   Call GetAsyncKeyState(vbKeyTab)
   Set clsSitMerc = New ControlaGrid
   Set clsSitMerc.ConexaoGrid = ADO_Cn_CD
   
   clsSitMerc.NomeFormulario = Me.Name
   clsSitMerc.NomeGrid = "grdItemNota"
   clsSitMerc.Colunas = 7
   clsSitMerc.LinhasVisiveis = 12
   
   clsSitMerc.Cabecalho = "<Referência; <Descrição; ^Quantidade; <Nota Fiscal; <Série; " _
                        & "<Data Entrada; ^Situação"
   
   clsSitMerc.Campos = "CI_Referencia; PR_Descricao; CI_Quantidade; CI_NotaFiscal; CI_Serie; CI_DataEntrada; CI_Situacao"
   
   clsSitMerc.Formato = "Caractere; Caractere; Numero; Numero; Caractere; Data; Caractere"
   
   clsSitMerc.Alinhamento = "Esquerda; Esquerda; Direita; Direita; Esuqerda; Esquerda; Esquerda"
   
   clsSitMerc.Tamanho = "1050; 5000; 1100; 1100; 560; 1300; 500"
   
   clsSitMerc.MontaCabecalho
   
End Sub



Private Sub optPesquisa1_Click(Index As Integer)
CheckOrdenaDescricao.Enabled = False
CheckOrdenadaReferencia.Enabled = False
End Sub

Private Sub optPesquisa2_Click(Index As Integer)
   If optPesquisa2(0).Value = True Then
      CheckOrdenaDescricao.Visible = False
      CheckOrdenadaReferencia.Visible = False
      txtPesquisa.Width = 4680
      txtPesquisa.MaxLength = 2
      CheckOrdenaDescricao.Enabled = False
      CheckOrdenadaReferencia.Enabled = False
   ElseIf optPesquisa2(1).Value = True Then
      CheckOrdenaDescricao.Visible = False
      CheckOrdenadaReferencia.Visible = False
      txtPesquisa.Width = 4680
      txtPesquisa.MaxLength = 7
      CheckOrdenaDescricao.Enabled = False
      CheckOrdenadaReferencia.Enabled = False
   ElseIf optPesquisa2(2).Value = True Then
      CheckOrdenaDescricao.Visible = True
      CheckOrdenadaReferencia.Visible = True
      txtPesquisa.Width = 550
      txtPesquisa.MaxLength = 4
      CheckOrdenaDescricao.Enabled = True
      CheckOrdenadaReferencia.Enabled = True
   End If

End Sub

Private Sub txtPesquisa_KeyPress(KeyAscii As Integer)
   VerTecla KeyAscii

End Sub

Private Sub txtPesquisa_LostFocus()
   If GetAsyncKeyState(vbKeyTab) <> 0 Then
      If CheckOrdenaDescricao.Visible = True And optPesquisa2(2).Value = True Then
         CheckOrdenaDescricao.SetFocus
      Else
         cmdPesquisa.SetFocus
      End If
   End If
End Sub
 
