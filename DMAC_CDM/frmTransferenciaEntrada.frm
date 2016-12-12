VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Begin VB.Form frmTransferenciaEntrada 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Transferência de Entrada"
   ClientHeight    =   8235
   ClientLeft      =   960
   ClientTop       =   2175
   ClientWidth     =   15780
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   15780
   ShowInTaskbar   =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   150
      TabIndex        =   9
      Top             =   150
      Width           =   4650
      Begin VB.OptionButton optTransferenciaProcessada 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Processada"
         ForeColor       =   &H00FFFFFF&
         Height          =   200
         Left            =   2175
         TabIndex        =   11
         Top             =   420
         Width           =   2190
      End
      Begin VB.OptionButton optTranferenciaAberta 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Aberta"
         ForeColor       =   &H00FFFFFF&
         Height          =   200
         Left            =   150
         TabIndex        =   10
         Top             =   420
         Value           =   -1  'True
         Width           =   1860
      End
      Begin VB.Label lblTrasnfAberto 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Transferência:"
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
         TabIndex        =   12
         Top             =   120
         Width           =   1935
      End
   End
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
   Begin VB.CommandButton cmdImprimir_1 
      Caption         =   "&Imprimir"
      Enabled         =   0   'False
      Height          =   375
      Left            =   12105
      TabIndex        =   5
      Top             =   10770
      Width           =   1230
   End
   Begin VB.TextBox txtSerie 
      BackColor       =   &H00A3A3A3&
      Height          =   300
      Left            =   8220
      TabIndex        =   3
      Top             =   10665
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.TextBox txtNotaFiscal 
      BackColor       =   &H00A3A3A3&
      Height          =   300
      Left            =   9315
      TabIndex        =   2
      Top             =   10665
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton cmdRetornar_1 
      Caption         =   "&Retornar"
      Height          =   375
      Left            =   13335
      TabIndex        =   1
      Top             =   10770
      Width           =   1155
   End
   Begin VB.CommandButton cmdOK_1 
      Caption         =   "&Processar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   10875
      TabIndex        =   0
      Top             =   10770
      Width           =   1230
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdDados 
      Height          =   5580
      Left            =   8400
      TabIndex        =   6
      Top             =   1080
      Width           =   6645
      _cx             =   11721
      _cy             =   9842
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
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmTransferenciaEntrada.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   2
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
   Begin VSFlex7DAOCtl.VSFlexGrid grdTransfAbertas 
      Height          =   5580
      Left            =   150
      TabIndex        =   7
      Top             =   1095
      Width           =   8130
      _cx             =   14340
      _cy             =   9842
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
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmTransferenciaEntrada.frx":00E2
      ScrollTrack     =   0   'False
      ScrollBars      =   2
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
   Begin CentroDeDistribuicao.chameleonButton cmdRetornar 
      Height          =   510
      Left            =   13620
      TabIndex        =   13
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
      MICON           =   "frmTransferenciaEntrada.frx":0199
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdImprimir 
      Height          =   510
      Left            =   12180
      TabIndex        =   14
      Top             =   6945
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Imprimir"
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
      MICON           =   "frmTransferenciaEntrada.frx":01B5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdOK 
      Height          =   510
      Left            =   10740
      TabIndex        =   15
      Top             =   6945
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Processar"
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
      MICON           =   "frmTransferenciaEntrada.frx":01D1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdConfirma 
      Height          =   510
      Left            =   9300
      TabIndex        =   16
      Top             =   6945
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Confirmar Recebimento"
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
      MICON           =   "frmTransferenciaEntrada.frx":01ED
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblDataNota 
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Data"
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   6765
      TabIndex        =   4
      Top             =   10665
      Visible         =   0   'False
      Width           =   1365
   End
End
Attribute VB_Name = "frmTransferenciaEntrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Conexao As New ADODB.Connection
Dim rdoDados As New ADODB.Recordset
Dim RdodadosItens As New ADODB.Recordset
Dim rsMovEstoque As New ADODB.Recordset
Dim rsDataControle As New ADODB.Recordset
Dim ControleLoja As New ADODB.Recordset
'Dim Servidor As String
Dim WBANCO As String
Dim sql As String
Dim ConverteQtde As Integer
Dim SomaTransferenciaEntrada As Integer
Dim CalculaCheckListEntrada As Long
Dim CalculaEstoqueFinal As Integer
Dim CalculaEstoque As Integer
Dim Loja As String
Dim wObservacao As String
Dim msgProcessamento As String

Dim RdoVar As New ADODB.Recordset


Function ConectaODBC(ByRef RdoVar, ByVal Usuario As String, ByVal Senha As String) As Boolean
'
'        On Error GoTo ConexaoErro
'
'        With RdoVar
'
'            'Servidor = GLB_Servidor
'            'Servidor = "DEMEO1"
'            WBANCO = GLB_Banco
'
'            .Connect = "Driver={SQL Server};" _
'                    & "Server=" & Trim(Servidor) & ";" _
'                    & "DataBase=" & Trim(WBANCO) & ";" _
'                    & "MaxBufferSize=512;" _
'                    & "PageTimeout=5;"
'                    '& "UID=" & Usuario & ";" _
'                    '& "PWD=" & Senha & ";"
'
'            .LoginTimeout = 5
'            .CursorDriver = rdUseClientBatch
'            '.EstablishConnection rdDriverNoPrompt
'        End With
'
'    If ConexaoDLLRDO.abrirConexaoRDODinamica(RdoVar) Then
'        ConectaODBC = True
'        Exit Function
'    End If
'
'ConexaoErro:
'
'    ConectaODBC = False



On Error GoTo ConexaoErro:
     
    If RdoVar.State = 1 Then
         RdoVar.Close
    End If
     
    '-- NOVA --
    If ConexaoDLLaDO.abrirConexaoADO(RdoVar, Nomeservidor, BancoDeDados) Then
        GLB_ConectouOK = True
        ConectaODBC = True
        Exit Function
    End If
     
    
ConexaoErro:
    MsgBox "Erro ao abrir banco de localizacao! "
       
    Exit Function

End Function
    
Private Sub cmbLojas_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        txtNotaFiscal.SetFocus
    End If
    
End Sub

Private Sub cmdNovaPesquisa_Click()

    Call Limpa

End Sub

Private Sub cmdConfirma_Click()
    cmdOK_Click
End Sub

Private Sub cmdImprimir_Click()
    Dim wContadorPagina As Integer
    Dim rdoCodBarra As New ADODB.Recordset
    
    'Declarações de variaveis para impressão
    Dim cabeEmpresa As String * 71
    Dim cabePagina As String * 8
    Dim cabeNota As String * 19
    Dim cabeRelatorio As String * 60
    Dim cabeSerie As String * 12
    Dim cabeLoja As String * 15
    Dim cabeDataHora As String * 36
    Dim cabeLinha As String * 80
    Dim cabeCabecalho As String * 80
    Dim detaReferencia As String * 13
    Dim detaCodBarra As String * 18
    Dim detaDescricao As String * 45
    Dim detaQtde As String * 4
    
    Screen.MousePointer = 11
    
    'For Each nomeImpressora In Printers
        'If Trim(nomeImpressora.DeviceName) = UCase(GLB_ImpCotacao) Then
            ' Seta impressora no sistema
            'Set Printer = nomeImpressora
            'Exit For
        'End If
    'Next
    
    cabeEmpresa = "DE MEO"
    cabePagina = "PAGINA: "
    cabeRelatorio = "                     CONFERENCIA DE TRANSFERENCIA DE ENTRADA"
    cabeNota = "NOTA FISCAL: " & grdTransfAbertas.TextMatrix(grdTransfAbertas.row, 1)
    cabeSerie = "   Serie: " & grdTransfAbertas.TextMatrix(grdTransfAbertas.row, 2)
    cabeLoja = "   DA LOJA: " & grdTransfAbertas.TextMatrix(grdTransfAbertas.row, 0)
    cabeDataHora = Space(14) & Date & Space(2) & time
    cabeLinha = "__________________________________________________________________________________"
    cabeCabecalho = "REFERENCIA  CODIGO BARRAS     DESCRICAO                                     QTDE"
    
    
    wPagina = 0
    wContadorPagina = 99
    Printer.ScaleMode = vbMillimeters
    Printer.ForeColor = "0"
    Printer.FontSize = 10
    Printer.FontName = "draft 20cpi"
    Printer.FontName = "COURIER"
    Printer.FontSize = 11
    Printer.FontBold = False
    Printer.DrawWidth = 3
    For i = 1 To grdDados.Rows - 1
        If wContadorPagina > 56 Then
           If i > 2 Then
               Printer.NewPage
           End If
           wPagina = wPagina + 1
           Printer.Print cabeEmpresa & cabePagina & wPagina
           Printer.Print cabeRelatorio
           Printer.Print cabeNota & cabeSerie & cabeLoja & cabeDataHora
           Printer.Print cabeLinha
           Printer.Print cabeCabecalho
           wContadorPagina = 6
        End If
        sql = ""
        sql = "Select PRB_CodigoBarras From ProdutoBarras Where PRB_Referencia = '" & grdDados.TextMatrix(i, 1) & "' " _
            & "and PRB_TipoCodigo = 'B'"
        'Set rdoCodBarra = rdoCnLoja.OpenResultset(SQL)
        rdoCodBarra.CursorLocation = adUseClient
        rdoCodBarra.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
        
        
        If Not rdoCodBarra.EOF Then
            detaReferencia = grdDados.TextMatrix(i, 1)
            detaCodBarra = left(rdoCodBarra("PRB_CodigoBarras") & Space(16), 16)
            detaDescricao = left(grdDados.TextMatrix(i, 2) & Space(40), 40)
            detaQtde = right(Space(4) & grdDados.TextMatrix(i, 3), 4)
            
            Printer.Print detaReferencia & detaCodBarra & detaDescricao & detaQtde
            
            rdoCodBarra.MoveNext
            If Not rdoCodBarra.EOF Then
              Do While Not rdoCodBarra.EOF
                  Printer.Print Space(13) & rdoCodBarra("prb_codigobarras")
                  rdoCodBarra.MoveNext
              Loop
            End If
        Else
            detaReferencia = grdDados.TextMatrix(i, 1)
            detaCodBarra = left(grdDados.TextMatrix(i, 1) & Space(16), 16)
            detaDescricao = left(grdDados.TextMatrix(i, 2) & Space(40), 40)
            detaQtde = right(Space(4) & grdDados.TextMatrix(i, 3), 4)
            
            Printer.Print detaReferencia & detaCodBarra & detaDescricao & detaQtde
        End If
        Printer.Print ""
        wContadorPagina = wContadorPagina + 2
        rdoCodBarra.Close
    Next i
    Printer.EndDoc
    
    Screen.MousePointer = 0

End Sub

Private Sub cmdOK_Click()
    Dim LojaDestino As String
    
    If grdDados.Rows > 1 Then
        Screen.MousePointer = 11
        'ProcessaNotaTransf
        cmdRetornar.SetFocus
        
        wObservacao = "OK - " & UCase(GLB_USU_Nome) & " - " & Format(Date, "dd/mm/yyyy") & " - " & Format(time, "hh:mm:ss")
        
        sql = "Exec SP_Est_Transferencia_destino '" & grdTransfAbertas.TextMatrix(grdTransfAbertas.row, 1) & "', '" _
        & grdTransfAbertas.TextMatrix(grdTransfAbertas.row, 2) & "', '" _
        & grdTransfAbertas.TextMatrix(grdTransfAbertas.row, 0) & "', '" _
        & wObservacao & "'"
        ADO_Cn_CD.Execute sql
        
        'SQL = "exec SP_Atualiza_Processos_Venda_Central"
        'ADO_Cn_CD_Loja.Execute SQL
        
        MsgBox "Nota " & Val(txtNotaFiscal.Text) & " processada com sucesso", vbInformation, "Sucesso"
        Call Limpa
        optTranferenciaAberta_Click
        
        Screen.MousePointer = 0
    Else
        MsgBox "Não Exite Nota Para Processar", vbCritical, "Atenção"
    End If
    
End Sub

Private Sub cmdRetornar_Click()
    
    Unload Me

End Sub
 

Private Sub cmdRetorna_Click()

End Sub

Private Sub Form_Activate()
    If GLB_modoOffline = True Then
        MsgBox "Você não tem conexão com a Retaguarda", vbExclamation, "Conexão Retaguarda"
        Unload Me
    Else
        Screen.MousePointer = 11
        optTranferenciaAberta_Click
        Screen.MousePointer = 0
    End If
End Sub

Private Sub Form_Load()
    
    carregarPosicaoTamanhoTela Me
    grdDados.ColWidth(6) = 0
    msgProcessamento = "Processada pelo Fechamento Mensal"

End Sub

Function AchaLoja() As String
    
    sql = ""
    sql = "Select CTS_Loja from ControleSistema"
    
    ControleLoja.CursorLocation = adUseClient
    ControleLoja.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    'rsDataControle.CursorLocation = adUseClient
    'rsDataControle.Open SQL, ADO_Cn_CD_Loja, adOpenForwardOnly, adLockPessimistic
    
    
    'Set ControleLoja = rdoCnLoja.OpenResultset("Select CT_Loja from Controle")
       
    AchaLoja = ControleLoja("CTS_Loja")
       
    ControleLoja.Close
   
End Function

Private Sub grdDadosX_Click()

End Sub



Private Sub grdTransfProcessadas_DblClick()
    
    Dim rdoDataCentral As New ADODB.Recordset
    Dim DataProc As Date
    
    sql = ""
    sql = "Select CT_DataAjuste from Controle"
        'Set rsDataControle = rdoCnLoja.OpenResultset(SQL)
        rsDataControle.CursorLocation = adUseClient
        rsDataControle.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
    
    sql = ""
    sql = "Select GetDate() as DataCentral"
    Set rdoDataCentral = ADO_Cn_CD.OpenResultset(sql)

    DataProc = rdoDataCentral("DataCentral")

    rdoDataCentral.Close
    
    If SelecionaCapaNfVenda(rdoDados, Format(DateAdd("d", -30, DataProc), "mm/dd/yyyy"), DataProc, "VC_Observacao not in ('0','') ") = True Then
        'grdTransfProcessadas.Rows = 1
        Do While Not rdoDados.EOF
            wObservacao = IIf(IsNull(Mid(rdoDados("VC_Observacao"), 1, 2)), "0", Mid(rdoDados("VC_Observacao"), 1, 2))
            If wObservacao <> "0" Then
                PreencheGrideNotaProcessada
            End If
            rdoDados.MoveNext
        Loop
    End If
    
End Sub

Private Sub Option1_Click()

End Sub

Private Sub Option2_Click()
    cmdOK.Enabled = False
    cmdImprimir.Enabled = False
End Sub

Public Function validaClickGrid(ByRef GradeUsu) As Boolean
    
    validaClickGrid = False
    If GradeUsu.row >= GradeUsu.FixedRows And GradeUsu.row < GradeUsu.Rows Then
        validaClickGrid = True
    End If

End Function

Private Sub grdTransfAbertas_click()

If validaClickGrid(grdTransfAbertas) = True Then
    
        Dim L As Integer
        
        If grdTransfAbertas.TextMatrix(grdTransfAbertas.row, 5) = msgProcessamento Then
            cmdOK.Enabled = False
            cmdConfirma.Enabled = True
        Else
            cmdOK.Enabled = True
            cmdConfirma.Enabled = False
        End If
        
    '
    ' perguntar se a linha do grid está em vermelho, se estiver proceder como nas processadas
    '
        Screen.MousePointer = 11
        If grdTransfAbertas.CellBackColor = &HFFFF80 Then    'Azul
            PreenheGrideItens grdTransfAbertas.TextMatrix(grdTransfAbertas.row, 1), grdTransfAbertas.TextMatrix(grdTransfAbertas.row, 2), grdTransfAbertas.TextMatrix(grdTransfAbertas.row, 0)
            'For L = 1 To grdTransfProcessadas.Rows - 1
                'grdTransfProcessadas.Row = L
                'If grdTransfProcessadas.CellBackColor = &HC0FFFF Then
                    'grdTransfProcessadas.Col = 0
                    'grdTransfProcessadas.ColSel = 3
                    'grdTransfProcessadas.FillStyle = flexFillRepeat
                    'grdTransfProcessadas.CellBackColor = &H80000005  'BRANCO
                    'grdTransfProcessadas.FillStyle = flexFillSingle
                'End If
            'Next L
            For L = 1 To grdTransfAbertas.Rows - 1
                grdTransfAbertas.row = L
                If grdTransfAbertas.CellBackColor = &HC0FFFF Then
                    grdTransfAbertas.col = 0
                    grdTransfAbertas.ColSel = 4
                    grdTransfAbertas.FillStyle = flexFillRepeat
                    grdTransfAbertas.CellBackColor = &H80000005  'BRANCO
                    grdTransfAbertas.FillStyle = flexFillSingle
                End If
            Next L
            cmdOK.Enabled = False
            cmdImprimir.Enabled = False
        Else
            PreenheGrideItens grdTransfAbertas.TextMatrix(grdTransfAbertas.row, 1), grdTransfAbertas.TextMatrix(grdTransfAbertas.row, 2), grdTransfAbertas.TextMatrix(grdTransfAbertas.row, 0)
            'PintaGrideNotaAbertas grdTransfAbertas.Row
            If grdTransfAbertas.Rows > 1 Then
                txtNotaFiscal.Text = grdTransfAbertas.TextMatrix(grdTransfAbertas.row, 1)
                txtSerie.Text = grdTransfAbertas.TextMatrix(grdTransfAbertas.row, 2)
                lblDataNota.Caption = grdTransfAbertas.TextMatrix(grdTransfAbertas.row, 3)
                Loja = grdTransfAbertas.TextMatrix(grdTransfAbertas.row, 0)
            End If
            'cmdOK.Enabled = True
            'cmdImprimir.Enabled = True
        End If
        Screen.MousePointer = 0
        
    Else
        limpaGrid grdDados
    End If


End Sub

Private Sub optTranferenciaAberta_Click()
    
    Screen.MousePointer = 11
    
    PreencheGrideNotaAberta
     
    If grdTransfAbertas.FixedRows < grdTransfAbertas.Rows Then
        cmdImprimir.Enabled = True
        cmdOK.Enabled = True
    Else
        cmdImprimir.Enabled = False
        cmdOK.Enabled = False
    End If
    
    Screen.MousePointer = 0
        
End Sub

Private Sub optTransferenciaProcessada_Click()

    Screen.MousePointer = 11

    If grdTransfAbertas.FixedRows < grdTransfAbertas.Rows Then
        cmdImprimir.Enabled = True
    Else
        cmdImprimir.Enabled = False
    End If
    
    PreencheGrideNotaProcessada
    
    cmdOK.Enabled = False
    
    Screen.MousePointer = 0
    
End Sub

Private Sub txtNotaFiscal_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 Then
        If IsNumeric(txtNotaFiscal.Text) = False Then
            txtNotaFiscal.SetFocus
            txtNotaFiscal.SelStart = 0
            txtNotaFiscal.SelLength = Len(txtNotaFiscal.Text)
        Else
            txtSerie.SetFocus
        End If
    End If
    
End Sub


Private Sub txtSerie_LostFocus()
    txtSerie.Text = UCase(txtSerie.Text)
End Sub

Sub Limpa()

    txtNotaFiscal.Text = ""
    txtSerie.Text = ""
    lblDataNota.Caption = ""
    lblDataNota.Visible = False
    grdDados.Rows = grdDados.FixedRows
    'grdDados.Rows = 2
    cmdOK.Enabled = False
    
End Sub

Function SelecionaCapaNfVenda(ByRef RdoVar, ByVal DataInicio As String, ByVal DataFim As String, ByVal Opcao As String) As Boolean

    'SQL = ""
    sql = "Select VC_DataEmissao,VC_TotalNota,VC_NotaFiscal,VC_Serie,VC_LojaOrigem,VC_Observacao from CapaNfVenda " _
        & "where VC_DataEmissao between '" & Format(DataInicio, "yyyy/mm/dd") & "' and '" & Format(DataFim, "yyyy/mm/dd") & "' " _
        & "and VC_TipoNota='T' and VC_LojaDestino='" & Trim(AchaLoja) & "' and " & Opcao _
        & " order by VC_LojaOrigem,VC_NotaFiscal"
    'Set RdoVar = ADO_Cn_CD.OpenResultset(SQL)
    RdoVar.CursorLocation = adUseClient
    RdoVar.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
    
    If Not RdoVar.EOF Then
       SelecionaCapaNfVenda = True
    Else
       SelecionaCapaNfVenda = False
    End If

End Function

Function SelecionaItemNfVenda(ByRef RdoVar, ByVal NotaFiscal As Double, ByVal serie As String, ByVal LojaOrigem As String) As Boolean
    
    sql = "Select VI_NotaFiscal,VI_Serie,vi_referencia,vi_quantidade, " _
    & "vi_precounitario,Vi_numeroitem,vi_valormercadoria,VI_TipoAtualizaTransito,pr_descricao " _
    & "from itemnfvenda,produto " _
    & "Where  VI_LojaOrigem = '" & LojaOrigem & "' " _
    & "and vI_notafiscal=" & NotaFiscal & " " _
    & "and vi_serie='" & serie & "' " _
    & "and vi_referencia=pr_referencia and vi_tiponota = 'T' order by Vi_numeroitem"
    'Set RdoVar = ADO_Cn_CD.OpenResultset(SQL)
    RdodadosItens.CursorLocation = adUseClient
    RdodadosItens.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
    
    If Not RdodadosItens.EOF Then
        SelecionaItemNfVenda = True
    Else
        SelecionaItemNfVenda = False
    End If
    
    'RdodadosItens.Close

End Function


Function PreencheGrideNotaAberta()

    Dim rdoDataCentral As New ADODB.Recordset
    Dim DataProc As Date
    Dim observacao As String
    Dim mesAtual As String
    Dim anoAtual As String
    Dim mesNota As String
    Dim anoNota As String
        
    limpaGrid grdTransfAbertas
        
    sql = ""
    sql = "Select GetDate() as DataCentral"
    rdoDataCentral.CursorLocation = adUseClient
    rdoDataCentral.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
    'Set rdoDataCentral = ADO_Cn_CD.OpenResultset(SQL)

    DataProc = rdoDataCentral("DataCentral")

    mesAtual = Format(DataProc, "MM")
    anoAtual = Format(DataProc, "YYYY")

    rdoDataCentral.Close

    If SelecionaCapaNfVenda(rdoDados, DateAdd("d", -30, DataProc), DataProc, "VC_Observacao in ('0') ") = True Then
        grdTransfAbertas.Rows = 1
        Do While Not rdoDados.EOF

                mesNota = Format(rdoDados("VC_DataEmissao"), "mm")
                anoNota = Format(rdoDados("VC_DataEmissao"), "yyyy")

                If anoAtual = anoNota And mesAtual = mesNota Then
                    observacao = ""
                Else
                    observacao = msgProcessamento
                End If

                grdTransfAbertas.AddItem Trim(rdoDados("VC_LojaOrigem")) & Chr(9) _
                    & rdoDados("VC_NotaFiscal") & Chr(9) _
                    & rdoDados("VC_Serie") & Chr(9) _
                    & Format(rdoDados("VC_DataEmissao"), "dd/mm/yyyy") & Chr(9) _
                    & Format(rdoDados("VC_TotalNota"), "##,###,###0.00") & Chr(9) _
                    & observacao
                    '& Format(rdoDados("VC_TotalNota"), "##,###,###0.00") & Chr(9) _
                    '& rdoDados("VC_Observacao")
                    
            'Else
                'grdTransfAbertas.AddItem Trim(rdoDados("VC_LojaOrigem")) & Chr(9) _
                    '& rdoDados("VC_NotaFiscal") & Chr(9) _
                    '& rdoDados("VC_Serie") & Chr(9) _
                    '& Format(rdoDados("VC_DataEmissao"), "dd/mm/yyyy") & Chr(9) _
                    '& Format(rdoDados("VC_TotalNota"), "##,###,###0.00")
            'End If
            rdoDados.MoveNext
        Loop
    End If
    
    rdoDados.Close
    limpaGrid grdDados
    
End Function


Function PreencheGrideNotaProcessada()

    Dim rdoDataCentral As New ADODB.Recordset
    Dim DataProc As Date
    
    'SQL = ""
    'SQL = "Select CTS_DataAjuste from ControleSistema"
    'rsDataControle.CursorLocation = adUseClient
    'rsDataControle.Open SQL, ADO_Cn_CD_Loja, adOpenForwardOnly, adLockPessimistic

    limpaGrid grdTransfAbertas

    sql = ""
    sql = "Select GetDate() as DataCentral"
    rdoDataCentral.CursorLocation = adUseClient
    rdoDataCentral.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    'Set rdoDataCentral = ADO_Cn_CD.OpenResultset(SQL)
        
    DataProc = rdoDataCentral("DataCentral")

    rdoDataCentral.Close

    If SelecionaCapaNfVenda(rdoDados, DateAdd("d", -30, DataProc), DataProc, "VC_Observacao <> '0' ") = True Then
        grdTransfAbertas.Rows = 1
        Do While Not rdoDados.EOF
            'wObservacao = IIf(IsNull(Mid(rdoDados("VC_Observacao"), 1, 2)), "0", Mid(rdoDados("VC_Observacao"), 1, 2))
            'If wObservacao = "0" Then
                grdTransfAbertas.AddItem Trim(rdoDados("VC_LojaOrigem")) & Chr(9) _
                    & rdoDados("VC_NotaFiscal") & Chr(9) _
                    & rdoDados("VC_Serie") & Chr(9) _
                    & Format(rdoDados("VC_DataEmissao"), "dd/mm/yyyy") & Chr(9) _
                    & Format(rdoDados("VC_TotalNota"), "##,###,###0.00") & Chr(9) _
                    & rdoDados("VC_Observacao")
            'Else
                'grdTransfAbertas.AddItem Trim(rdoDados("VC_LojaOrigem")) & Chr(9) _
                    '& rdoDados("VC_NotaFiscal") & Chr(9) _
                    '& rdoDados("VC_Serie") & Chr(9) _
                    '& Format(rdoDados("VC_DataEmissao"), "dd/mm/yyyy") & Chr(9) _
                    '& Format(rdoDados("VC_TotalNota"), "##,###,###0.00")
            'End If
            rdoDados.MoveNext
        Loop
    End If
    
    limpaGrid grdDados
    rdoDados.Close
    
End Function


'Function PreencheGrideNotaProcessada()

    'grdTransfProcessas.AddItem Trim(rdoDados("VC_LojaOrigem")) & Chr(9) _
        & rdoDados("VC_NotaFiscal") & Chr(9) _
        & rdoDados("VC_Serie") & Chr(9) _
        & rdoDados("VC_Observacao")
    
'End Function

Function PreenheGrideItens(ByVal NotaFiscal As Double, ByVal serie As String, ByVal LojaOrigem As String)

    If SelecionaItemNfVenda(RdodadosItens, NotaFiscal, serie, LojaOrigem) = True Then
        grdDados.Rows = 1
        Do While Not RdodadosItens.EOF
            grdDados.AddItem RdodadosItens("Vi_numeroItem") & Chr(vbKeyTab) & RdodadosItens("Vi_Referencia") _
                & Chr(vbKeyTab) & RdodadosItens("Pr_Descricao") _
                & Chr(vbKeyTab) & RdodadosItens("Vi_Quantidade") _
                & Chr(vbKeyTab) & Format(RdodadosItens("VI_PrecoUnitario"), "##,###,###0.00") _
                & Chr(vbKeyTab) & Format(RdodadosItens("VI_ValorMercadoria"), "##,###,###0.00") _
                & Chr(vbKeyTab) & IIf(IsNull(RdodadosItens("VI_TipoAtualizaTransito")), 0, RdodadosItens("VI_TipoAtualizaTransito"))
            RdodadosItens.MoveNext
        Loop
    End If

    RdodadosItens.Close

End Function


Function PintaGrideNotaAbertas(ByVal Linha As Integer)
    Dim L As Integer
    
    'For L = 1 To grdTransfProcessadas.Rows - 1
        'grdTransfProcessadas.Row = L
        'If grdTransfProcessadas.CellBackColor = &HC0FFFF Then
            'grdTransfProcessadas.Col = 0
            'grdTransfProcessadas.ColSel = 3
            'grdTransfProcessadas.FillStyle = flexFillRepeat
            'grdTransfProcessadas.CellBackColor = &H80000005  'BRANCO
            'grdTransfProcessadas.FillStyle = flexFillSingle
        'End If
    'Next L
    If grdTransfAbertas.CellBackColor = &HC0FFFF Then
        grdTransfAbertas.col = 0
        grdTransfAbertas.ColSel = 4
        grdTransfAbertas.FillStyle = flexFillRepeat
        grdTransfAbertas.CellBackColor = &H80000005  'BRANCO
        grdTransfAbertas.FillStyle = flexFillSingle
        grdDados.Rows = 1
    Else
        For L = 1 To grdTransfAbertas.Rows - 1
            grdTransfAbertas.row = L
            If grdTransfAbertas.CellBackColor = &HC0FFFF Then
                grdTransfAbertas.col = 0
                grdTransfAbertas.ColSel = 4
                grdTransfAbertas.FillStyle = flexFillRepeat
                grdTransfAbertas.CellBackColor = &H80000005  'BRANCO
                grdTransfAbertas.FillStyle = flexFillSingle
            End If
        Next L
        
        grdTransfAbertas.row = Linha
        grdTransfAbertas.col = 0
        grdTransfAbertas.ColSel = 4
        grdTransfAbertas.FillStyle = flexFillRepeat
        grdTransfAbertas.CellBackColor = &HC0FFFF  'AMARELOGRID
        grdTransfAbertas.FillStyle = flexFillSingle
   End If
   
End Function

Sub PintaGrideNotaProcessadas(ByVal Linha As Integer)
    Dim L As Integer
    
    For L = 1 To grdTransfAbertas.Rows - 1
        grdTransfAbertas.row = L
        If grdTransfAbertas.CellBackColor = &HC0FFFF Then
            grdTransfAbertas.col = 0
            grdTransfAbertas.ColSel = 4
            grdTransfAbertas.FillStyle = flexFillRepeat
            grdTransfAbertas.CellBackColor = &H80000005  'BRANCO
            grdTransfAbertas.FillStyle = flexFillSingle
        End If
    Next L
    'If grdTransfProcessadas.CellBackColor = &HC0FFFF Then
        'grdTransfProcessadas.Col = 0
        'grdTransfProcessadas.ColSel = 3
        'grdTransfProcessadas.FillStyle = flexFillRepeat
        'grdTransfProcessadas.CellBackColor = &H80000005  'BRANCO
        'grdTransfProcessadas.FillStyle = flexFillSingle
        'grdDados.Rows = 1
    'Else
        'For L = 1 To grdTransfProcessadas.Rows - 1
            'grdTransfProcessadas.Row = L
            'If grdTransfProcessadas.CellBackColor = &HC0FFFF Then
                'grdTransfProcessadas.Col = 0
                'grdTransfProcessadas.ColSel = 3
                'grdTransfProcessadas.FillStyle = flexFillRepeat
                'grdTransfProcessadas.CellBackColor = &H80000005  'BRANCO
                'grdTransfProcessadas.FillStyle = flexFillSingle
            'End If
        'Next L
        
        'grdTransfProcessadas.Row = Linha
        'grdTransfProcessadas.Col = 0
        'grdTransfProcessadas.ColSel = 3
        'grdTransfProcessadas.FillStyle = flexFillRepeat
        'grdTransfProcessadas.CellBackColor = &HC0FFFF  'AMARELOGRID
        'grdTransfProcessadas.FillStyle = flexFillSingle
   'End If
   
End Sub

Function ProcessaNotaTransf()
    
    Dim LojaDestino As String
    Dim wObservacao As String
    
    For i = 1 To grdDados.Rows - 1
        Screen.MousePointer = 11
        LojaDestino = AchaLoja
            
        If grdDados.TextMatrix(i, 6) = 0 Then
            
            'MovimentoEstoqueCentral LojaDestino, Date, grdDados.TextMatrix(i, 1)
            
            '
            '-----------------Atualiza Estoque e Transito Central
            '
            
            sql = "Update Estoque set es_estoque = es_estoque + " & grdDados.TextMatrix(i, 3) & "," _
                & "es_transito = es_transito - " & grdDados.TextMatrix(i, 3) & "" _
                & " where es_referencia = '" & grdDados.TextMatrix(i, 1) & "' " _
                & " and es_loja = '" & LojaDestino & "'"

            ADO_Cn_CD.Execute (sql)
            
'            grdDados.TextMatrix(i, 6) = 1
            
            '
            '-----------------Atualiza Tipo de Atualizacao do item
            '
            sql = ""
            sql = "Update ItemNfVenda set VI_TipoAtualizaTransito=1 " _
                & "Where VI_NotaFiscal = " & Val(txtNotaFiscal.Text) & " and " _
                & "VI_Serie = '" & txtSerie.Text & "' and VI_LojaOrigem = '" & Loja & "' and " _
                & "VI_DataEmissao = '" & Format(lblDataNota.Caption, "mm/dd/yyyy") & "'" _
                & "and VI_Referencia='" & grdDados.TextMatrix(i, 1) & "'"
                ADO_Cn_CD.Execute (sql)
                
        End If
            
        '
        '---------------------Processando Loja
        '
        
        If grdDados.TextMatrix(i, 6) = 1 Then
        
            'SQL = "Select * from estoque, MovimentacaoEstoque  " _
                & "where ME_Referencia='" & grdDados.TextMatrix(i, 1) & "' and ME_Loja='" & LojaDestino & "' " _
                & "and ME_DataMovimento='" & Format(Date, "mm/dd/yyyy") & "'"
            'Set rsMovEstoque = rdoCnLoja.OpenResultset(SQL)
            
            'If rsMovEstoque.EOF Then
'Voltar:
 '               SQL = "Select * from estoque " _
                    & "where es_referencia = '" & grdDados.TextMatrix(i, 1) & "' "
                'Set RsdadosItens = rdoCnLoja.OpenResultset(SQL)
                'If Not RsdadosItens.EOF Then
                '     ConverteQtde = grdDados.TextMatrix(i, 3)
                '     CalculaEstoqueFinal = RsdadosItens("es_Estoque") + ConverteQtde
                '     CalculaCheckListEntrada = Val(txtNotaFiscal.Text)
                 
                '     rdoCnLoja.Execute "Insert into MovimentacaoEstoque (ME_DataMovimento, ME_Loja, ME_Referencia, " _
                       & "ME_EstoqueInicial, ME_Venda, ME_TransferenciaSaida, ME_SRS, ME_AjusteSaida, " _
                       & "ME_DevolucaoCompras, ME_SRE, ME_TranferenciaEntrada, ME_AjusteEntrada, " _
                       & "ME_DevolucaoVenda, ME_EstoqueFinal, ME_ChekListEntrada, ME_ChekListSaida, " _
                       & "ME_Situacao, ME_MovimentoOK) Values ('" & Format(Date, "mm/dd/yyyy") & "', " _
                       & "'" & LojaDestino & "', '" & grdDados.TextMatrix(i, 1) & "'," & RsdadosItens("es_Estoque") & "," _
                       & "0,0,0,0,0,0," & grdDados.TextMatrix(i, 3) & ",0,0," & CalculaEstoqueFinal & "," & CalculaCheckListEntrada & " ,0,0 ,'S')"
                 'Else
'            If MsgBox("Referencia " & grdDados.TextMatrix(i, 1) & " não encontrada, deseja cadastrar agora e continuar o processo", vbYesNo + vbInformation, "Aviso") = vbYes Then
'                GravaProduto grdDados.TextMatrix(i, 1)
'            Else
'                MsgBox "Não foi possivel finalizar o processo de entrada desta nota", vbCritical, "Atenção"
'                Exit Function
'            End If
'            Else
'                SQL = "Select * from estoque " _
'                    & "where es_referencia = '" & grdDados.TextMatrix(i, 1) & "' "
'                Set RsdadosItens = rdoCnLoja.OpenResultset(SQL)
'                If Not RsdadosItens.EOF Then
'                   ConverteQtde = grdDados.TextMatrix(i, 3)
'                   SomaTransferenciaEntrada = rsMovEstoque("ME_TranferenciaEntrada") + ConverteQtde
'                   CalculaCheckListEntrada = rsMovEstoque("ME_ChekListEntrada") + txtNotaFiscal.Text
'                   CalculaEstoqueFinal = RsdadosItens("es_Estoque") + ConverteQtde
'
'                   SQL = "Update MovimentacaoEstoque set ME_TranferenciaEntrada = " & SomaTransferenciaEntrada & ", " _
'                                & "ME_EstoqueFinal=" & CalculaEstoqueFinal & ", ME_ChekListEntrada= " & CalculaCheckListEntrada & " " _
'                                & "where ME_DataMovimento= '" & Format(Date, "mm/dd/yyyy") & "' and ME_Loja= '" & LojaDestino & "' and " _
'                                & "ME_Referencia='" & grdDados.TextMatrix(i, 1) & "' "
'                   rdoCnLoja.Execute (SQL)
'                End If
'            End If
'            CalculaEstoque = RsdadosItens("es_Estoque") + ConverteQtde
'            ConverteQtde = grdDados.TextMatrix(i, 3) * (-1)
'            'CalculaEstoque = RsdadosItens("es_Estoque") + ConverteQtde
'
'            rdoCnLoja.Execute "Update estoque set es_Estoque= " & CalculaEstoque & " where es_Referencia='" & grdDados.TextMatrix(i, 1) & "' and es_Loja='" & LojaDestino & "' "
'
'            rdoCnLoja.Execute "Insert into EstqLojaDBF (Referencia,Quantidade,Situacao) Values ('" & grdDados.TextMatrix(i, 1) & "'," & ConverteQtde & ",'A')"
            
            
            sql = "Select es_Referencia from estoque " _
                & "where es_referencia = '" & grdDados.TextMatrix(i, 1) & "' "
            'Set RsdadosItens = rdoCnLoja.OpenResultset(SQL)
            RsdadosItens.CursorLocation = adUseClient
            RsdadosItens.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
            
            
            If RsdadosItens.EOF Then
                'If MsgBox("Referencia " & grdDados.TextMatrix(i, 1) & " não encontrada, deseja cadastrar agora e continuar o processo", vbYesNo + vbInformation, "Aviso") = vbYes Then
                    'GravaProduto grdDados.TextMatrix(i, 1)
                'Else
                    MsgBox "Não foi possivel finalizar o processo de entrada desta nota, existe referência não cadastrada", vbCritical, "Atenção"
                    RsdadosItens.Close
                    Exit Function
                'End If
            End If
            
            RsdadosItens.Close
            'GravaControleEstoque grdDados.TextMatrix(i, 1), grdDados.TextMatrix(i, 3), "T", txtNotaFiscal.Text
            sql = ""
            sql = "Update ItemNfVenda set VI_TipoAtualizaTransito=1 " _
                & "Where VI_NotaFiscal = " & Val(txtNotaFiscal.Text) & " and " _
                & "VI_Serie = '" & txtSerie.Text & "' and VI_LojaOrigem = '" & Loja & "' and " _
                & "VI_DataEmissao = '" & Format(lblDataNota.Caption, "mm/dd/yyyy") & "'" _
                & "and VI_Referencia = '" & grdDados.TextMatrix(i, 1) & "'"
                ADO_Cn_CD.Execute (sql)
        End If
    Next
    '
    '----------------------Atualizando Capa Central
    '
    wObservacao = "OK - " & UCase(GLB_USU_Nome) & " - " & Format(Date, "dd/mm/yy") & " - " & Format(time, "hh:mm:ss")
        
    sql = "Update CapaNFVenda Set VC_Observacao = '" & wObservacao & "', " _
        & "VC_TipoAtualizaTransito= 1, VC_DataTransferenciaEntrada='" & Format(Date, "mm/dd/yyyy") & "' " _
        & "Where VC_NotaFiscal = " & Val(txtNotaFiscal.Text) & " and " _
        & "VC_Serie = '" & txtSerie.Text & "' and VC_LojaOrigem = '" & Loja & "' and " _
        & "VC_DataEmissao = '" & Format(lblDataNota.Caption, "mm/dd/yyyy") & "'"
    ADO_Cn_CD.Execute (sql)
    
    'SQL = "update GLB_ControleTarefas set CTA_Transferencia = 'S'"
    'ADO_Cn_CD.Execute (SQL)
       
    grdTransfAbertas.col = 0
    grdTransfAbertas.ColSel = 4
    grdTransfAbertas.FillStyle = flexFillRepeat
    grdTransfAbertas.CellBackColor = &HFFFF80     'Azul
    grdTransfAbertas.FillStyle = flexFillSingle
       
       
End Function


Function VerificaUsuarioSenha(ByVal Usuario As String, ByVal Senha As String) As Boolean
    Dim rsVerSenha As New ADODB.Recordset
    
    sql = ""
    sql = "Select Us_TipoUsuario from Usuario " _
        & "where US_TipoUsuario in (4,2) " _
        & "and US_Usuario = '" & GLB_USU_Nome & "' " _
        & "and US_Senha='" & Senha & "'"
        'Set rsVerSenha = rdoCnLoja.OpenResultset(SQL)
        rsVerSenha.CursorLocation = adUseClient
        rsVerSenha.Open sql, ADO_Cn_CD, adOpenForwardOnly, adLockPessimistic
        
    If Not rsVerSenha.EOF Then
        VerificaUsuarioSenha = True
    Else
        VerificaUsuarioSenha = False
    End If
    
End Function

