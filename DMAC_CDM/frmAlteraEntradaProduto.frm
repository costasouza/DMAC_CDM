VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsflex7.ocx"
Begin VB.Form frmAlteraEntrdaProduto 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Alterar Quantida de Produto Nota"
   ClientHeight    =   7860
   ClientLeft      =   3795
   ClientTop       =   1530
   ClientWidth     =   15420
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   15420
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   44
      Left            =   150
      ScaleHeight     =   45
      ScaleWidth      =   14880
      TabIndex        =   16
      Top             =   6810
      Width           =   14880
   End
   Begin VB.Frame frameItens 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   150
      TabIndex        =   9
      Top             =   5865
      Width           =   14910
      Begin VB.TextBox txtReferencia 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   120
         MaxLength       =   20
         TabIndex        =   10
         Tag             =   "Refêrencia / Código de Barras / Código Produto Fornecedor"
         Top             =   370
         Width           =   4440
      End
      Begin VB.TextBox txtQuantidade 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   4650
         MaxLength       =   4
         TabIndex        =   11
         ToolTipText     =   "Quantidade"
         Top             =   375
         Width           =   660
      End
      Begin VB.TextBox txtPrecoUnitario 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   5400
         MaxLength       =   8
         TabIndex        =   12
         ToolTipText     =   "Preço Unitário"
         Top             =   370
         Width           =   1260
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Refêrencia / Código de Barras / Código Produto Fornecedor"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "Código Produto"
         Top             =   120
         Width           =   4275
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde."
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   4660
         TabIndex        =   14
         Top             =   120
         Width           =   390
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preço Unitário"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   5420
         TabIndex        =   13
         Top             =   120
         Width           =   1005
      End
   End
   Begin VB.Frame frmEntrada 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   14910
      Begin VB.TextBox txtNomeFantasia 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   945
         MaxLength       =   60
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Fornecedor"
         Top             =   350
         Width           =   2940
      End
      Begin VB.TextBox txtSerie 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   5460
         MaxLength       =   2
         TabIndex        =   4
         ToolTipText     =   "Série"
         Top             =   350
         Width           =   400
      End
      Begin VB.TextBox txtNotaFiscal 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   3975
         MaxLength       =   8
         TabIndex        =   3
         ToolTipText     =   "Nota Fiscal"
         Top             =   350
         Width           =   1400
      End
      Begin VB.TextBox txtCodigoFornecedor 
         BackColor       =   &H00C0C0C0&
         Height          =   315
         Left            =   120
         MaxLength       =   4
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Código Fornecedor"
         Top             =   350
         Width           =   735
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo        Fornecedor"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   1665
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Série"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   5460
         TabIndex        =   6
         Top             =   120
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nota Fiscal"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   3975
         TabIndex        =   5
         Top             =   120
         Width           =   795
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grdPrincipal 
      Height          =   4635
      Left            =   150
      TabIndex        =   8
      Top             =   1065
      Width           =   14910
      _cx             =   26300
      _cy             =   8176
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
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmAlteraEntradaProduto.frx":0000
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
      ShowComboButton =   0   'False
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   0
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   5263440
      ForeColorFrozen =   4210752
      WallPaperAlignment=   9
   End
   Begin CentroDeDistribuicao.chameleonButton cmdRetorna 
      Height          =   510
      Left            =   13620
      TabIndex        =   17
      ToolTipText     =   "Grava"
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
      MICON           =   "frmAlteraEntradaProduto.frx":0166
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdLimpa 
      Height          =   510
      Left            =   12180
      TabIndex        =   18
      ToolTipText     =   "Grava"
      Top             =   6945
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Limpar"
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
      MICON           =   "frmAlteraEntradaProduto.frx":0182
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
Attribute VB_Name = "frmAlteraEntrdaProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rdoFornecedor As New ADODB.Recordset
Dim rdoItemNota As New ADODB.Recordset
Dim sql As String
Dim antigoQtd As Integer
Dim qtnNova As Integer
Dim novoCusto As Double
Dim antigoCusto As Double
Dim valorNovo As Double
Dim codFor As String



        
Private Sub cmdLimpa_Click()
Call limpaTudo
End Sub

Private Sub cmdRetorna_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
txtCodigoFornecedor.SetFocus

   
   carregarPosicaoTamanhoTela Me
End Sub

Private Sub grdPrincipal_DblClick()
If grdPrincipal.row > 0 Then

codFor = Trim(grdPrincipal.TextMatrix(grdPrincipal.row, 0))
txtReferencia.Text = grdPrincipal.TextMatrix(grdPrincipal.row, 2)
txtQuantidade.Text = grdPrincipal.TextMatrix(grdPrincipal.row, 3)
antigoQtd = grdPrincipal.TextMatrix(grdPrincipal.row, 3)
txtPrecoUnitario.Text = grdPrincipal.TextMatrix(grdPrincipal.row, 4)
novoCusto = grdPrincipal.TextMatrix(grdPrincipal.row, 9)
antigoCusto = grdPrincipal.TextMatrix(grdPrincipal.row, 10)
txtQuantidade.SetFocus
End If
End Sub



Private Sub txtCodigoFornecedor_Click()
Call limpaTudo
End Sub

Private Sub txtCodigoFornecedor_KeyPress(KeyAscii As Integer)


Select Case KeyAscii
     Case vbKeyDelete
     Case vbKeyBack
     Case 48 To 57
     Case 13
     Case Else
             Beep
             KeyAscii = 0
End Select


    If KeyAscii = 13 Then
     Call carregarfornecedor
End If

End Sub


Private Sub txtCodigoFornecedor_LostFocus()
If txtCodigoFornecedor.Text <> "" Then

Call carregarfornecedor
End If

End Sub

Private Sub txtPrecoUnitario_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

 Call gravarNovo
End If

End Sub

Private Sub txtQuantidade_KeyPress(KeyAscii As Integer)
Dim novo As Double
    If (KeyAscii = 13) Then
      
    novo = novoPreco(txtQuantidade.Text, antigoQtd, txtPrecoUnitario.Text)
    txtPrecoUnitario.Text = Format(novo, "##,#0.00")
    valorNovo = txtPrecoUnitario.Text
    
    novo = 0
    novoCusto = novoPreco(txtQuantidade.Text, antigoQtd, novoCusto)
    
    
    
        
    novo = 0
    antigoCusto = novoPreco(txtQuantidade.Text, antigoQtd, novoCusto)
    
    
    
    qtnNova = txtQuantidade.Text
    
    
    
    txtPrecoUnitario.SetFocus
    End If

End Sub

Private Sub txtSerie_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call carregarProduto
End If
End Sub


Private Function carregarDescricaoProduto(referencia As String) As String
    Dim rdoDescricao As New ADODB.Recordset
    Dim sql As String
    
    sql = "select pr_descricao from produto where pr_referencia = '" _
    & referencia & "'"
    
    rdoDescricao.CursorLocation = adUseClient
    rdoDescricao.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If rdoDescricao.BOF And rdoDescricao.EOF Then
        carregarDescricaoProduto = " (Não foi possível encontra o produto) "
    Else
        carregarDescricaoProduto = rdoDescricao("pr_descricao")
    End If
    
    rdoDescricao.Close
End Function



Private Function carregarProduto()
   grdPrincipal.Rows = 1
    If txtNomeFantasia.Text <> "" And txtCodigoFornecedor.Text <> "" And txtNotaFiscal.Text <> "" And txtSerie.Text <> "" Then
    sql = " Select CI_PrecoUnitario, CI_CustoAnterior, CI_quantidade,CI_CFOP, CI_NovoCusto,CI_ValorICMSST,CI_ValorIPI,CI_ValorICMS,CI_DescricaoFornecedor,CI_codigoProduto,CI_Referencia from  Itemnfcompra where  ci_notafiscal=" & txtNotaFiscal.Text & " and ci_serie='" & txtSerie.Text & "' and CI_Fornecedor =" & txtCodigoFornecedor.Text & ""
    
    rdoItemNota.CursorLocation = adUseClient
    rdoItemNota.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If rdoItemNota.EOF = False Then
        
       Do While Not rdoItemNota.EOF

       
          
        grdPrincipal.AddItem rdoItemNota("CI_codigoProduto") & Chr(9) & rdoItemNota("CI_Referencia") _
        & Chr(9) & rdoItemNota("CI_DescricaoFornecedor") & Chr(9) _
        & rdoItemNota("CI_quantidade") & Chr(9) & Format(rdoItemNota("CI_PrecoUnitario"), "##,#0.00") _
        & Chr(9) & rdoItemNota("CI_CFOP") & Chr(9) & Format(rdoItemNota("CI_ValorICMS"), "##,#0.00") & Chr(9) _
        & Format(rdoItemNota("CI_ValorIPI"), "##,#0.00") & Chr(9) & Format(rdoItemNota("CI_ValorICMSST"), "##,#0.00") & Chr(9) _
         & Format(rdoItemNota("CI_NovoCusto"), "##,#0.00") & Chr(9) & Format(rdoItemNota("CI_CustoAnterior"), "##,#0.00")
       
       
       rdoItemNota.MoveNext
       
       
       Loop
       
       rdoItemNota.Close
    End If
    
    
    End If

End Function

Private Function carregarfornecedor()

   sql = "select FO_NomeFantasia from fornecedor where FO_codigoFornecedor = '" _
        & txtCodigoFornecedor.Text & "'"
        
    rdoFornecedor.CursorLocation = adUseClient
    rdoFornecedor.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If rdoFornecedor.BOF And rdoFornecedor.EOF Then
      MsgBox ("Fornecedor Não  Encontrado!")
        
    Else
        txtNomeFantasia.Text = rdoFornecedor("FO_NomeFantasia")

         txtNotaFiscal.SetFocus
    End If
  rdoFornecedor.Close
    

End Function


Private Function novoPreco(novaQtd As Integer, antigoQtd As Integer, Preco As Double) As Double

Dim novo As Double

    novo = Preco * antigoQtd
    novo = novo / novaQtd
    novoPreco = novo
   

End Function

Private Sub txtSerie_LostFocus()
Call carregarProduto
End Sub


Private Function limpaTudo()
grdPrincipal.Rows = 1
        txtReferencia.Text = ""
        txtQuantidade.Text = ""
        txtPrecoUnitario.Text = ""
        grdPrincipal.Rows = 1
        txtCodigoFornecedor.Text = ""
        txtNomeFantasia.Text = ""
        txtNotaFiscal.Text = ""
        txtSerie.Text = ""
        sql = ""
        antigoQtd = 0
        qtnNova = 0
        novoCusto = 0
        antigoCusto = 0
        codFor = ""
        valorNovo = 0
End Function

Private Function gravarNovo()
sql = ""
 If (qtnNova > 0 And valorNovo <> 0) Then
                
                
                    sql = " update Itemnfcompra  set   CI_PrecoUnitario=" & ConverteVirgula(valorNovo) & ", CI_CustoAnterior=" & ConverteVirgula(antigoCusto) & ", CI_quantidade=" & qtnNova & ", CI_NovoCusto=" & ConverteVirgula(novoCusto) & "" _
                    & " where  ci_notafiscal=" & txtNotaFiscal.Text & " and ci_serie='" & txtSerie.Text & "' and CI_Fornecedor =" & txtCodigoFornecedor.Text & " and CI_codigoProduto='" & codFor & "'"
                   
                   
                ADO_Cn_CDLocal.BeginTrans
                ADO_Cn_CDLocal.Execute (sql)
                ADO_Cn_CDLocal.CommitTrans
                
                
                
        antigoQtd = 0
        qtnNova = 0
        novoCusto = 0
        antigoCusto = 0
        codFor = ""
        valorNovo = 0
        txtReferencia.Text = ""
        txtQuantidade.Text = ""
        txtPrecoUnitario.Text = ""
End If
Call carregarProduto

End Function
