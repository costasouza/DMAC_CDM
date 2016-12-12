VERSION 5.00
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Begin VB.Form frmControleCD 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Controle Dmac"
   ClientHeight    =   11685
   ClientLeft      =   1635
   ClientTop       =   510
   ClientWidth     =   14820
   Icon            =   "frmControleCD.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmControleCD.frx":23FA
   ScaleHeight     =   11685
   ScaleWidth      =   14820
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox cmdAgenda 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2220
      Left            =   4650
      Picture         =   "frmControleCD.frx":B78B
      ScaleHeight     =   2220
      ScaleWidth      =   2220
      TabIndex        =   1
      Top             =   210
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VSFlex7DAOCtl.VSFlexGrid grdHistorico 
      Height          =   5655
      Left            =   360
      TabIndex        =   3
      Top             =   2940
      Visible         =   0   'False
      Width           =   10350
      _cx             =   18256
      _cy             =   9975
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   0
      ForeColor       =   12632256
      BackColorFixed  =   0
      ForeColorFixed  =   0
      BackColorSel    =   0
      ForeColorSel    =   12632256
      BackColorBkg    =   0
      BackColorAlternate=   0
      GridColor       =   0
      GridColorFixed  =   0
      TreeColor       =   0
      FloodColor      =   5263440
      SheetBorder     =   0
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
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmControleCD.frx":12132
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
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Label lblModoOff 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "T R A B A L H A N D O   N O   M O D O   O F F - L I N E"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   90
      TabIndex        =   2
      Top             =   2580
      Width           =   15180
   End
   Begin VB.Image cmdVolta 
      Height          =   480
      Left            =   16575
      Picture         =   "frmControleCD.frx":121CC
      Top             =   10815
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image cmdAvanca 
      Height          =   480
      Left            =   17175
      Picture         =   "frmControleCD.frx":1285D
      Top             =   10815
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image cmdHome 
      Height          =   480
      Left            =   15165
      Picture         =   "frmControleCD.frx":12EF3
      Top             =   10815
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image cmdSair 
      Height          =   480
      Left            =   14500
      Picture         =   "frmControleCD.frx":1368F
      Top             =   10815
      Width           =   480
   End
   Begin VB.Image cmdBotao 
      Height          =   465
      Index           =   11
      Left            =   13695
      Top             =   10815
      Width           =   435
   End
   Begin VB.Image cmdBotao 
      Height          =   465
      Index           =   10
      Left            =   12900
      Top             =   10815
      Width           =   435
   End
   Begin VB.Image cmdBotao 
      Height          =   465
      Index           =   9
      Left            =   12100
      Top             =   10815
      Width           =   435
   End
   Begin VB.Image cmdBotao 
      Height          =   465
      Index           =   8
      Left            =   11300
      Top             =   10815
      Width           =   435
   End
   Begin VB.Image cmdBotao 
      Height          =   465
      Index           =   7
      Left            =   10500
      Top             =   10815
      Width           =   435
   End
   Begin VB.Image cmdBotao 
      Height          =   465
      Index           =   6
      Left            =   4390
      Top             =   10815
      Width           =   435
   End
   Begin VB.Image cmdBotao 
      Height          =   465
      Index           =   5
      Left            =   3590
      Top             =   10815
      Width           =   435
   End
   Begin VB.Image cmdBotao 
      Height          =   465
      Index           =   4
      Left            =   2790
      Top             =   10815
      Width           =   435
   End
   Begin VB.Image cmdBotao 
      Height          =   465
      Index           =   3
      Left            =   1995
      Top             =   10815
      Width           =   435
   End
   Begin VB.Image cmdBotao 
      Height          =   465
      Index           =   2
      Left            =   1185
      Top             =   10815
      Width           =   435
   End
   Begin VB.Image cmdBotao 
      Height          =   465
      Index           =   1
      Left            =   420
      Top             =   10815
      Width           =   435
   End
   Begin VB.Image imgVersaoDMAC 
      Height          =   675
      Left            =   6765
      Picture         =   "frmControleCD.frx":13D29
      Top             =   10680
      Width           =   1935
   End
   Begin VB.Image webPadraoTamanhoJanela 
      Height          =   7590
      Left            =   90
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   15180
   End
   Begin VB.Label lblNomeTelas 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label lblNomeTelas"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FA9923&
      Height          =   300
      Left            =   135
      TabIndex        =   0
      Top             =   2565
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.Image Image1 
      Height          =   2475
      Left            =   90
      MouseIcon       =   "frmControleCD.frx":157C3
      Picture         =   "frmControleCD.frx":15ACD
      Stretch         =   -1  'True
      Top             =   100
      Width           =   15180
   End
End
Attribute VB_Name = "frmControleCD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cont As Integer
Dim nomeBotoes As String
Dim tamanhoVetorBotoes As Byte
Dim posicaoInicio As Byte

Dim interlavoBotao As Byte
Dim exibirXML As Boolean

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdAgenda_Click()
        frmAgendaDeRecebimento.Show 1
        frmControleCD.lblNomeTelas.Visible = True
        frmControleCD.lblNomeTelas.Caption = frmAgendaDeRecebimento.Caption
End Sub



Private Sub cmdAgenda_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'cmdAgenda.Visible = True
    'lblteste.Caption = "x = " & X & vbNewLine & "y = " & Y
End Sub

Private Sub cmdBotao_Click(Index As Integer)
    cmdHome.Visible = True
    ChamaTelaMenu cmdBotao(Index).ToolTipText
End Sub

Private Sub cmdBotao_MouseOver(Index As Integer)
    'cmdBotao(Index).ForeOver = &HFCC785
End Sub



Private Sub cmdHome_Click()
    cmdHome.Visible = False
    wCont = 1
    menuProximo "000000", False
    'lblPosicao.Caption = ""
    'lblPosicaoSub.Caption = ""
    lblNomeTelas.Caption = ""
    'Image1.Picture = LoadPicture("C:\Sistemas\cd\Imagens\t3.JPG")
    ' Image4.Visible = False
End Sub

Private Sub cmdAvanca_Click()
    wControle = wAuxMenuControle & wAuxMenuControle2
    menuProximo wControle, True
End Sub

Private Sub cmdRecebimento_Click(Index As Integer)
End Sub

Private Sub cmdvolta_Click()
    menuProximo wControle, False
End Sub

Private Sub cmdSair_Click()
    Call AlterarResolucao(resolucaoOriginal.Colunas, resolucaoOriginal.Linhas)
    sairDoSistema
End Sub


Private Sub cmdVersaoDmac_Click()
    
End Sub

Private Sub CarregaHistorico()
    Dim sql As String
    Dim adoHistorico As New ADODB.Recordset
    
    sql = "select top 10 nf,serie,dataemi as data,CN_DescricaoOperacao as descricao,TIPONOTA from nfcapa,codigooperacaonovo where CN_CodigoOperacaoNovo = CODOPER and tiponota in ('V','T','S','E','C') order by nf desc"
    
    adoHistorico.CursorLocation = adUseClient
    adoHistorico.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
    grdHistorico.AddItem ""
    
    Do While Not adoHistorico.EOF
        grdHistorico.AddItem adoHistorico("nf") & Chr(9) & adoHistorico("serie") & Chr(9) & adoHistorico("data") & Chr(9) & adoHistorico("descricao") & Chr(9) & adoHistorico("TIPONOTA")
        adoHistorico.MoveNext
    Loop
    
    adoHistorico.Close


End Sub

Private Sub Form_Activate()
    verificarAtualizacaoAplicativo
    If buscaAutomaticaXML = True And exibirXML = True Then
        exibirXML = False
        frmXML.Show 1
    End If
End Sub

Private Sub Form_Load()

    cmdAgenda.left = 4650
    cmdAgenda.top = 225

    exibirXML = True
        
    If Not ConectaADO() Then
        msnon = "Trabalhando  Off-line"
        lblModoOff.Visible = True
        GLB_modoOffline = True
    Else
       msnon = ""
       lblModoOff.Visible = False
       GLB_modoOffline = False
    End If
    
    'CarregaHistorico

    acertaTamanhoBotoes
    acertaTamanhoVetorBotoes cmdBotao
    'acertaTamanhoVetorBotoes cmdBotoesConsulta
    'acertaTamanhoVetorBotoes cmdBotoesOperacao

    wCont = 1
    menuProximo "000000", False

    
    If verificaPermissao("frmAgendaDeRecebimento") Then
        cmdAgenda.Enabled = True
    Else
        cmdAgenda.Enabled = False
    End If
    
    posicaoTelaY = webPadraoTamanhoJanela.top
    posicaoTelaX = webPadraoTamanhoJanela.left
    tamanhoTelaY = webPadraoTamanhoJanela.Height
    tamanhoTelaX = webPadraoTamanhoJanela.Width
    
    Image1.Picture = LoadPicture(endIMG("topo1024768hd"))
End Sub



Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    'lblteste.Caption = "x = " & X & vbNewLine & "y = " & Y
    If (x > 4440 And x < 7000) And (y > 240 And y < 2000) Then
        cmdAgenda.Visible = True
    Else
        cmdAgenda.Visible = False
    End If
End Sub

Private Sub Image1_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    cmdAgenda.Visible = False
End Sub

Private Sub imgVersaoDMAC_Click()
    MsgBox "Logado como: " & GLB_USU_Nome & vbNewLine & "DMAC CDM Versão: " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub acertaTamanhoVetorBotoes(botoes)
    For i = 1 To botoes.Count - 1
        botoes(i).Width = 425
        botoes(i).Height = 465
        botoes(i).top = 10815
    Next i
End Sub

Private Sub acertaTamanhoBotoes()

    cmdHome.top = 10815
    'cmdAvanca.top = 10815
    'cmdVolta.top = 10815
    cmdSair.top = 10815
    
    cmdHome.Height = 465
    'cmdAvanca.Height = 465
    'cmdVolta.Height = 465
    cmdSair.Height = 465
    
    cmdHome.Width = 435
    'cmdAvanca.Width = 435
    'cmdVolta.Width = 435
    cmdSair.Width = 435

    cmdSair.left = 14500
    cmdHome.left = 14500
    
    'cmdVolta.left = cmdBotao(1).left
    'cmdAvanca.left = cmdBotao(cmdBotao.UBound).left
    
End Sub

Private Sub carregaPosicaoBotoes(botao)

    tamanhoVetorBotoes = botao.Count - 1
    
    If tamanhoVetorBotoes > 8 Then
        cmdAvanca.Visible = True
        cmdVolta.Visible = True
        botao(0).left = posicaoInicio + interlavoBotao
    Else
        botao(0).left = posicaoInicio
    End If
    
    botao(0).Visible = True
    
    For i = 1 To tamanhoVetorBotoes
        botao(i).left = botao(i - 1).left + interlavoBotao
        botao(i).Visible = True
        If limiteCentro(botao(i).left) Then
            botao(i).left = 10500
        End If
    Next i
    
    'cmdHome.Visible = True
    
End Sub

Function limiteCentro(i As Integer) As Boolean
    limiteCentro = False
End Function

    Public Sub verificarAtualizacaoAplicativo()
        ShellExecute hwnd, "open", ("C:\Sistemas\DMAC CDM\TrocaVersao.exe"), "", "", 1
    End Sub




