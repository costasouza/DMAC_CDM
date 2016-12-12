VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Begin VB.Form frmTransferenciaSemEstoque 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Transferência de Mercadoria sem Estoque "
   ClientHeight    =   7890
   ClientLeft      =   1140
   ClientTop       =   3555
   ClientWidth     =   15105
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   7890
   ScaleWidth      =   15105
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "\"
      Height          =   6570
      Left            =   10395
      TabIndex        =   9
      Top             =   75
      Width           =   4740
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   45
      Left            =   75
      ScaleHeight     =   45
      ScaleWidth      =   15045
      TabIndex        =   7
      Top             =   6780
      Width           =   15045
   End
   Begin VB.Frame fraPesquisa 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Pesquisa"
      ForeColor       =   &H00FF0000&
      Height          =   930
      Left            =   90
      TabIndex        =   6
      Top             =   75
      Width           =   10245
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
         Left            =   4005
         TabIndex        =   0
         Top             =   435
         Width           =   5970
      End
      Begin VB.OptionButton optPesquisa 
         BackColor       =   &H00404040&
         Caption         =   "Referência"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   3
         Top             =   510
         Value           =   -1  'True
         Width           =   1080
      End
      Begin VB.OptionButton optPesquisa 
         BackColor       =   &H00404040&
         Caption         =   "Fornecedor"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   1
         Left            =   1470
         TabIndex        =   4
         Top             =   510
         Width           =   1110
      End
      Begin VB.OptionButton optPesquisa 
         BackColor       =   &H00404040&
         Caption         =   "Descrição"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   2
         Left            =   2805
         TabIndex        =   5
         Top             =   510
         Width           =   1020
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
         Left            =   75
         TabIndex        =   10
         Top             =   90
         Width           =   825
      End
   End
   Begin CentroDeDistribuicao.chameleonButton cmdPesquisa 
      Height          =   510
      Left            =   12300
      TabIndex        =   1
      Top             =   6900
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
      MICON           =   "frmTransferenciaSemEstoque.frx":0000
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
      Left            =   13725
      TabIndex        =   2
      Top             =   6900
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
      MICON           =   "frmTransferenciaSemEstoque.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdItens 
      Height          =   5550
      Left            =   75
      TabIndex        =   8
      Top             =   1080
      Width           =   10260
      _cx             =   18098
      _cy             =   9790
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
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmTransferenciaSemEstoque.frx":0038
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
Attribute VB_Name = "frmTransferenciaSemEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim adoPesquisa As New ADODB.Recordset
Dim adoLoja As New ADODB.Recordset

Dim wWhere As String
Dim ControlaGrid As String

Dim i As Integer

Dim wQuanti As Double

Private Sub cmdPesquisa_Click()

If txtPesquisa.Text = "" Then
    MsgBox "Preencha o campo de pesquisa!", vbCritical, "ATENÇÃO"
    txtPesquisa.SetFocus
    Exit Sub
End If

Screen.MousePointer = 11
If optPesquisa(0).Value = True = True And txtPesquisa.Text <> "" Then
   
    wWhere = " Es_Referencia = '" & txtPesquisa.Text & "'"
End If

If optPesquisa(1).Value = True And txtPesquisa.Text <> "" Then
    
     wWhere = " Pr_CodigoFornecedor = '" & txtPesquisa.Text & "'"
End If

If optPesquisa(2).Value = True And txtPesquisa.Text <> "" Then
    
    wWhere = " PR_Descricao  like '" & Mid(txtPesquisa.Text, 1, 3) & "%' "
End If

grdItens.Rows = 1

Call CompletaGrid

Screen.MousePointer = 0

End Sub

Private Sub cmdRetorna_Click()
    frmControleCD.lblNomeTelas.Caption = ""
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    carregarPosicaoTamanhoTela Me
    'JanelaTOP Me
    
grdItens.MergeRow(0) = True

For i = 0 To 4

grdItens.MergeCol(i) = True

Next i

End Sub
Private Sub grdItens_EnterCell()

If grdItens.col = 3 Or grdItens.col = 4 Then
    
    grdItens.EditCell

End If

End Sub
Private Sub grdItens_KeyPress(KeyAscii As Integer)

wQuanti = grdItens.TextMatrix(grdItens.row, 3)

If wQuanti <= 0 Then
   
    MsgBox "Quantidade inferior à permitida", vbInformation
    grdItens.TextMatrix(grdItens.row, 3) = ""
    grdItens.EditCell
    grdItens.col = 3
    
    
Else
               
   sql = "Select Lo_Loja, * from Loja where Lo_Loja = '" & grdItens.TextMatrix(grdItens.row, 4) & "' " _
        & " and Lo_Loja not in ('CMC','CD','CMCS','CMCE','182','183','184','CONSO')"
        
     adoLoja.CursorLocation = adUseClient
     adoLoja.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
      
    

    If adoLoja.EOF Then
            
        MsgBox "Loja não cadastrada", vbInformation
        grdItens.TextMatrix(grdItens.row, 4) = ""
        grdItens.EditCell
        grdItens.col = 4
        
    Else
    
        ADO_Cn_CDLocal.BeginTrans
               
        sql = "Insert Into Romaneio (RO_NumeroRomaneio,RO_LojaOrigem,RO_LojaDestino," _
             & "RO_Referencia,RO_DataSolicitacao,RO_QuantidadePedida,RO_Tipo,RO_Situacao)" _
             & "values(0,'CD'," & grdItens.TextMatrix(grdItens.row, 4) & ",'" _
             & grdItens.TextMatrix(grdItens.row, 0) & "', Convert(Char(10), GetDate(), 101)," _
             & grdItens.TextMatrix(grdItens.row, 3) & ",'N','A')"
             
        ADO_Cn_CDLocal.Execute (sql)
        ADO_Cn_CDLocal.CommitTrans
        
    
    
        ADO_Cn_CDLocal.BeginTrans
               
        sql = "Update Estoque Set ES_Estoque = ES_Estoque - " _
            & grdItens.TextMatrix(grdItens.row, 3) & " where Es_Loja = 'CD' and " _
            & "Es_Referencia = '" & grdItens.TextMatrix(grdItens.row, 0) & "' "
             
        ADO_Cn_CDLocal.Execute (sql)
        ADO_Cn_CDLocal.CommitTrans
        
        
        ADO_Cn_CDLocal.BeginTrans
               
         sql = "Update Estoque Set ES_Romaneio = ES_Romaneio + " _
            & grdItens.TextMatrix(grdItens.row, 3) & " where Es_Loja = '" _
            & grdItens.TextMatrix(grdItens.row, 4) & "' and " _
            & "Es_Referencia = '" & grdItens.TextMatrix(grdItens.row, 0) & "' "
             
        ADO_Cn_CDLocal.Execute (sql)
        ADO_Cn_CDLocal.CommitTrans
    
        
        
        grdItens.Rows = 1
        
        Call CompletaGrid

    End If
     
End If

End Sub

Sub CompletaGrid()

sql = "Select es_loja,es_referencia,es_estoque,pr_descricao from estoque,produto" _
    & " where es_loja='cd' and es_referencia=pr_referencia and es_estoque < 1 " _
    & "and " & wWhere & "  and es_referencia not in(select ci_referencia " _
    & "from itemnfcompra,capanfcompra where cc_notafiscal=ci_notafiscal " _
    & "and cc_serie=ci_serie and cc_fornecedor=ci_fornecedor and cc_loja='cd' " _
    & "and ci_situacao ='E')"
    
    
    adoPesquisa.CursorLocation = adUseClient
    adoPesquisa.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic


If Not adoPesquisa.EOF Then

        Do While Not adoPesquisa.EOF

                grdItens.AddItem adoPesquisa("Es_Referencia") & Chr(9) & _
                adoPesquisa("Pr_Descricao") & Chr(9) & _
                adoPesquisa("Es_Estoque")

            adoPesquisa.MoveNext

        Loop
Else

        MsgBox "Nenhum registro foi encontrado para a informação desejada", vbInformation
        txtPesquisa.Text = ""
        txtPesquisa.SetFocus

End If
adoPesquisa.Close

End Sub

Private Sub optPesquisa_Click(Index As Integer)
    grdItens.Rows = 1
    txtPesquisa.Text = ""
    txtPesquisa.SetFocus
End Sub
