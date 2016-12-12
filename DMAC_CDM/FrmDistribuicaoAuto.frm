VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Begin VB.Form FrmDistribuicaoAuto 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Distribuição Automática"
   ClientHeight    =   7650
   ClientLeft      =   3735
   ClientTop       =   1545
   ClientWidth     =   15990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   15990
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   45
      Left            =   150
      ScaleHeight     =   45
      ScaleWidth      =   14910
      TabIndex        =   8
      Top             =   6810
      Width           =   14910
   End
   Begin VB.Frame FrmDistribuicao 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   840
      Left            =   150
      TabIndex        =   4
      Top             =   240
      Width           =   14880
      Begin VB.TextBox Txtnomeforne 
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
         Height          =   300
         Left            =   3900
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   375
         Width           =   6150
      End
      Begin VB.TextBox Txtfornecedor 
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
         Height          =   300
         Left            =   2835
         MaxLength       =   4
         TabIndex        =   2
         Top             =   375
         Width           =   945
      End
      Begin VB.TextBox Txtserie 
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
         Height          =   300
         Left            =   1770
         MaxLength       =   2
         TabIndex        =   1
         Top             =   375
         Width           =   960
      End
      Begin VB.TextBox TxtNota 
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
         Height          =   300
         Left            =   150
         MaxLength       =   10
         TabIndex        =   0
         Top             =   375
         Width           =   1455
      End
      Begin VB.Label lblQuantidadeDistribuida 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade Distribuida"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   11640
         TabIndex        =   17
         Top             =   120
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label lblQuantidadeDist 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   12885
         TabIndex        =   16
         Top             =   360
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblQuantidade 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   11055
         TabIndex        =   15
         Top             =   360
         Width           =   360
      End
      Begin VB.Label lblQuantidadeNota 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade Nota"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   10200
         TabIndex        =   14
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Lblfornecedor 
         BackColor       =   &H00404040&
         Caption         =   "Fornecedor"
         ForeColor       =   &H00C0C0C0&
         Height          =   210
         Left            =   2835
         TabIndex        =   7
         Top             =   105
         Width           =   1020
      End
      Begin VB.Label Lblserie 
         BackColor       =   &H00404040&
         Caption         =   "Série"
         ForeColor       =   &H00C0C0C0&
         Height          =   210
         Left            =   1785
         TabIndex        =   6
         Top             =   105
         Width           =   705
      End
      Begin VB.Label LblNota 
         BackColor       =   &H00404040&
         Caption         =   "Nota Fiscal"
         ForeColor       =   &H00C0C0C0&
         Height          =   210
         Left            =   150
         TabIndex        =   5
         Top             =   105
         Width           =   1140
      End
   End
   Begin CentroDeDistribuicao.chameleonButton cmdPesquisa 
      Height          =   510
      Left            =   12180
      TabIndex        =   9
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
      MICON           =   "FrmDistribuicaoAuto.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdRetornar 
      Height          =   510
      Left            =   13620
      TabIndex        =   10
      Top             =   6945
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Retornar"
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
      MICON           =   "FrmDistribuicaoAuto.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VSFlex7DAOCtl.VSFlexGrid GrdDistribuicao 
      Height          =   5355
      Left            =   150
      TabIndex        =   11
      Top             =   1200
      Width           =   14880
      _cx             =   26247
      _cy             =   9446
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
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmDistribuicaoAuto.frx":0038
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
   Begin CentroDeDistribuicao.chameleonButton cmdEstoque 
      Height          =   510
      Left            =   10740
      TabIndex        =   12
      Top             =   6945
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Estoque"
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
      MICON           =   "FrmDistribuicaoAuto.frx":0183
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdTransfere 
      Height          =   510
      Left            =   9300
      TabIndex        =   13
      Top             =   6945
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Romaneio Automático"
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
      MICON           =   "FrmDistribuicaoAuto.frx":019F
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
Attribute VB_Name = "FrmDistribuicaoAuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim adoDistribuicao As New ADODB.Recordset
Dim adoMinMax As New ADODB.Recordset

Dim wcoluna As Integer
Dim wcontrolacor As Integer

Dim cor As String
Dim cor1 As String
Dim cor2 As String

Private WithEvents clsGrddistribuicao As ControlaGrid
Attribute clsGrddistribuicao.VB_VarHelpID = -1

Private Sub cmdEstoque_Click()
    CriaRomaneio Trim(TxtNota.Text), Trim(txtFornecedor.Text), Trim(txtSerie.Text), "2"
End Sub

Private Sub cmdRetornar_Click()
    frmControleCD.lblNomeTelas.Caption = ""
    Unload Me
End Sub


Private Sub cmdTransfere_Click()
    CriaRomaneio Trim(TxtNota.Text), Trim(txtFornecedor.Text), Trim(txtSerie.Text), "1"
End Sub

Private Sub Form_Load()
    carregarPosicaoTamanhoTela Me
    cmdEstoque.Enabled = False
    cmdTransfere.Enabled = False
    'JanelaTOP Me
End Sub


Private Sub cmdPesquisa_Click()
    Dim i As Integer
    i = 0

    Screen.MousePointer = 11
    
    GrdDistribuicao.Rows = 1
    GrdDistribuicao.AddItem ""
    GrdDistribuicao.RemoveItem (1)
    
   ' Valida Campo Nota
    If TxtNota.Text = "" Then
        TxtNota.SetFocus
        TxtNota.SelStart = 0
        TxtNota.SelLength = Len(TxtNota.Text)
        MsgBox "Favor preencher o campo Nota Fiscal", vbInformation
        Screen.MousePointer = 0
        Exit Sub
        
    ' Valida Campo Série
    ElseIf txtSerie.Text = "" Then
        txtSerie.SetFocus
        txtSerie.SelStart = 0
        txtSerie.SelLength = Len(txtSerie.Text)
        MsgBox "Favor preencher o campo Série", vbInformation
        Screen.MousePointer = 0
        Exit Sub
        
    ' Valida Campo Fornecedor
    ElseIf txtFornecedor.Text = "" Then
        txtFornecedor.SetFocus
        txtFornecedor.SelStart = 0
        txtFornecedor.SelLength = Len(txtFornecedor.Text)
        MsgBox "Favor preencher o campo Fornecedor", vbInformation
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    
    
' Verifica se há algum registro para Distribuição Automática
    
    sql = "Select * from DistribuicaoAutomatica where da_distribuicaoAtendida <> 0 and da_notafiscal = '" & TxtNota.Text & "' and da_serie = '" & txtSerie.Text & "' and da_codigofornecedor = '" & txtFornecedor.Text & "' "

    adoDistribuicao.CursorLocation = adUseClient
    adoDistribuicao.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic

    If adoDistribuicao.EOF Then
       
        MsgBox "Não há necessidade de distribuição.", vbExclamation, "Atenção"

        TxtNota.Text = ""
        txtSerie.Text = ""
        txtFornecedor.Text = ""
        Txtnomeforne.Text = ""
        Screen.MousePointer = 0
        Exit Sub

    End If

    adoDistribuicao.Close
    
    sql = "select da_estoque,da_romaneio,da_transito,da_referencia,pr_descricao,da_lojadestino,da_codigofornecedor," _
        & "da_percentualdistribuicao,da_distribuicaocalculada," _
        & "da_distribuicaoatendida, da_situacao from Produto, DistribuicaoAutomatica,estoque" _
        & " where da_distribuicaoCalculada <> 0 and es_referencia = da_referencia  and es_loja = da_lojaDestino and da_notafiscal = " & TxtNota.Text _
        & " and da_serie = '" & txtSerie.Text & "' and da_codigofornecedor = '" _
        & txtFornecedor.Text & "'  and pr_referencia = da_referencia" _
        & " order by es_referencia,es_estoqueMaximo desc ,da_regiao desc"
    
    adoDistribuicao.CursorLocation = adUseClient
    adoDistribuicao.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic

    If Not adoDistribuicao.EOF Then
    
        Do While Not adoDistribuicao.EOF
            sql = "SELECT ES_EstoqueMinimo,ES_EstoqueMaximo,ES_Display, ES_MinimoInformado, ES_MaximoInformado " & _
                  "From Estoque Where ES_Referencia = '" & adoDistribuicao("da_referencia") & "' " & _
                  "and ES_Loja = '" & adoDistribuicao("DA_LojaDestino") & "'"
          
            adoMinMax.CursorLocation = adUseClient
            adoMinMax.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    

            If Not adoMinMax.EOF Then
                
                Totaliza GrdDistribuicao.Rows, adoDistribuicao("da_referencia")
                
                GrdDistribuicao.AddItem adoDistribuicao("da_referencia") & Chr(9) & _
                adoDistribuicao("PR_Descricao") & Chr(9) & _
                adoDistribuicao("DA_LojaDestino") & Chr(9) & _
                adoDistribuicao("DA_DistribuicaoAtendida") & Chr(9) & _
                IIf(IsNull(adoDistribuicao("DA_Estoque")), 0, adoDistribuicao("DA_Estoque")) & Chr(9) & _
                IIf(IsNull(adoDistribuicao("DA_Romaneio")), 0, adoDistribuicao("DA_Romaneio")) & Chr(9) & _
                IIf(IsNull(adoDistribuicao("DA_Transito")), 0, adoDistribuicao("DA_Transito")) & Chr(9) & _
                adoMinMax("ES_EstoqueMinimo") & Chr(9) & _
                adoMinMax("ES_EstoqueMaximo") & Chr(9) & _
                adoMinMax("ES_Display") & Chr(9) & _
                adoDistribuicao("DA_Situacao") & Chr(9)
            End If
            adoDistribuicao.MoveNext
            adoMinMax.Close
        Loop
       
        If GrdDistribuicao.Rows > 2 And GrdDistribuicao.TextMatrix(1, 0) = "" Then
           GrdDistribuicao.RemoveItem 1
        End If
        cmdEstoque.Enabled = True
        cmdTransfere.Enabled = True
        adoDistribuicao.Close
        Totaliza GrdDistribuicao.Rows, ""
        
        QuantidadeDistribuida TxtNota.Text, txtSerie.Text, txtFornecedor.Text, GrdDistribuicao.TextMatrix(GrdDistribuicao.row, 1)

  Else
     
     MsgBox "Registro não encontrado", vbInformation
     TxtNota.Text = ""
     TxtNota.SetFocus
     txtSerie.Text = ""
     txtFornecedor.Text = ""
     Txtnomeforne.Text = ""
     cmdEstoque.Enabled = False
     cmdTransfere.Enabled = False
      
  End If

Screen.MousePointer = 0

End Sub


Private Sub CriaDistribuicao(nota As String, serie As String, fornecedor As String, referencia As String, qtdeNF As String, qtdeCD As String)

    Dim sql As String
    
    sql = "Exec SP_Distribuicao_Automatica_Compras '" & nota & "', '" & serie & "', '" & fornecedor & "', '" & referencia & "', " & qtdeNF & ", " & qtdeCD & ", 99"
    
    ADO_Cn_CDLocal.Execute (sql)

End Sub

Private Sub GrdDistribuicao_AfterEdit(ByVal row As Long, ByVal col As Long)
    Dim quantidade As String
    Dim Loja As String
    Dim referencia As String
    
    If vbKeyReturn Then
        quantidade = GrdDistribuicao.TextMatrix(GrdDistribuicao.row, 3)
        Loja = GrdDistribuicao.TextMatrix(GrdDistribuicao.row, 2)
        referencia = GrdDistribuicao.TextMatrix(GrdDistribuicao.row, 0)
    
        AtualizaQuantidade quantidade, Loja, referencia
        QuantidadeDistribuida TxtNota.Text, txtSerie.Text, txtFornecedor.Text, GrdDistribuicao.TextMatrix(GrdDistribuicao.row, 0)
        cmdPesquisa_Click
    End If
End Sub


Private Sub GrdDistribuicao_EnterCell()
   If GrdDistribuicao.row > 1 Then
        QuantidadeDistribuida TxtNota.Text, txtSerie.Text, txtFornecedor.Text, GrdDistribuicao.TextMatrix(GrdDistribuicao.row, 0)
    End If
End Sub

Private Sub txtFornecedor_LostFocus()
    ' Cria Distribuição Automática
    
    Dim nota As String
    Dim serie As String
    Dim fornecedor As String
    Dim referencia As String
    Dim qtdeNF As String
    Dim qtdeCD As String
    Dim sql As String
    Dim ref As String
    Dim rsNota As New ADODB.Recordset
    Dim rsEstoque As New ADODB.Recordset
    
    
    If txtFornecedor.Text <> "" Then
        sql = "Select fo_razaosocial from fornecedor where fo_codigofornecedor = " & txtFornecedor & ""
        
        adoDistribuicao.CursorLocation = adUseClient
        adoDistribuicao.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
        If adoDistribuicao.EOF Then
           MsgBox "Nenhum registro encontrado", vbInformation
        Else
                Txtnomeforne.Text = adoDistribuicao("fo_razaosocial")
                nota = TxtNota.Text
                serie = txtSerie.Text
                fornecedor = txtFornecedor.Text
        End If
        adoDistribuicao.Close
                
        sql = "select da_referencia from distribuicaoAutomatica where da_NotaFiscal = '" & nota & "' and da_serie = '" & serie & "'"
    
        rsNota.CursorLocation = adUseClient
        
        rsNota.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
        If Not rsNota.EOF Then
            sql = "delete from distribuicaoAutomatica where da_NotaFiscal = '" & nota & "' and da_serie = '" & serie & "'"
            ADO_Cn_CDLocal.Execute (sql)
        End If
    
        rsNota.Close
    
        sql = " select ci_referencia, ci_quantidade from itemnfcompra where ci_notafiscal = '" & nota & "' and ci_serie = '" & serie & "' and ci_fornecedor like '" & fornecedor & "'"
    
        rsNota.CursorLocation = adUseClient
        rsNota.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
              
        Do While Not rsNota.EOF
            
            sql = " select es_estoque from estoque where es_referencia = '" & rsNota("ci_referencia") & "' and es_loja= 'CD'"
            
            rsEstoque.CursorLocation = adUseClient
            rsEstoque.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
            If Not rsEstoque.EOF Then
            
                If rsEstoque("es_estoque") >= rsNota("ci_quantidade") Then
            
                    CriaDistribuicao nota, serie, fornecedor, rsNota("ci_referencia"), rsNota("ci_quantidade"), rsNota("ci_quantidade")
            
                End If
            End If
                rsEstoque.Close
                rsNota.MoveNext
        Loop
                        
        TxtNota.Enabled = False
        txtSerie.Enabled = False
        txtFornecedor.Enabled = False
    End If

End Sub

Private Sub CriaRomaneio(nota As String, fornecedor As String, serie As String, operacao As String)
    
    Dim sql As String
    
    sql = "Exec SP_CriaRomaneio_Distribuicao_Automatica " & nota & ", '" & serie & "', " & fornecedor & ", " & operacao
    ADO_Cn_CDLocal.Execute (sql)
    
    GrdDistribuicao.Rows = 1
    GrdDistribuicao.AddItem ""
    GrdDistribuicao.RemoveItem (1)
    
End Sub

Private Sub AtualizaQuantidade(quantidade As String, Loja As String, referencia As String)
    
    Dim sql As String
    
    sql = "update DistribuicaoAutomatica set da_distribuicaoAtendida =  " & quantidade & _
          " where da_LojaDestino = '" & Loja & _
          "' and da_referencia = '" & referencia & "'"
    
    ADO_Cn_CDLocal.Execute sql
    
    
End Sub

Private Sub QuantidadeDistribuida(nota As String, serie As String, fornecedor As String, referencia As String)

    Dim sql As String
    Dim rsQuantidade As New ADODB.Recordset
    
    sql = "select sum(da_distribuicaoAtendida) as atendida, ci_quantidade" _
        & " from distribuicaoAutomatica, itemnfcompra where da_notafiscal = " & nota _
        & " and da_serie = '" & serie & "'" _
        & " and da_codigoFornecedor = " & fornecedor _
        & " and da_referencia = '" & Trim(referencia) & "'" _
        & " and ci_referencia = da_referencia and ci_notafiscal = da_notafiscal and ci_serie = da_serie and da_codigoFornecedor = ci_fornecedor" _
        & " group by ci_quantidade"
        rsQuantidade.CursorLocation = adUseClient
        rsQuantidade.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
        If Not rsQuantidade.EOF Then
            lblQuantidade.Caption = rsQuantidade("ci_quantidade")
            lblQuantidadeDist.Caption = rsQuantidade("atendida")
        End If
        rsQuantidade.Close
    
End Sub

Private Sub Totaliza(Linha As Integer, ReferenciaAtual As String)
    
    Dim ReferenciaAnterior As String
    Dim Total As String
    Dim sql As String
    Dim rsNota As New ADODB.Recordset
    
    If (Linha - 1) > 0 Then
        
        ReferenciaAnterior = GrdDistribuicao.TextMatrix((Linha - 1), 0)
        
        sql = "select sum(da_DistribuicaoAtendida) as quantidade from distribuicaoAutomatica where da_referencia = '" & ReferenciaAnterior & "'" _
        & " and da_notafiscal = " & TxtNota.Text _
        & " and da_serie = '" & txtSerie.Text & "'" _
        & " and da_codigoFornecedor = " & txtFornecedor.Text
        
        rsNota.CursorLocation = adUseClient
        rsNota.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
        If Not rsNota.EOF Then
            Total = rsNota("quantidade")
        End If
        
        rsNota.Close
        
        If Trim(ReferenciaAtual) <> Trim(ReferenciaAnterior) Then
            GrdDistribuicao.AddItem Chr(9) & "Total" & Chr(9) & Chr(9) & Total
            
            GrdDistribuicao.FillStyle = flexFillRepeat
            GrdDistribuicao.row = GrdDistribuicao.Rows - 1
            GrdDistribuicao.RowSel = GrdDistribuicao.Rows - 1
            GrdDistribuicao.col = 0
            GrdDistribuicao.ColSel = GrdDistribuicao.Cols - 1
            GrdDistribuicao.CellFontBold = True
            GrdDistribuicao.CellForeColor = &H8000000D
            GrdDistribuicao.FillStyle = flexFillSingle
        
        End If
    
    End If
    
End Sub

'22/09/2006
'Por: Celso
'Adicionado ao grid as colunas: estoque,romaneio,transito,minimo,maximo e display
Private Sub Txtserie_Change()

End Sub

Private Sub txtSerie_LostFocus()
    
    If txtSerie.Text <> "" Then
        txtSerie.Text = UCase(txtSerie.Text)
    End If

End Sub
