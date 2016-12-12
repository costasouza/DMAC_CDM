VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmXML 
   Appearance      =   0  'Flat
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Leitura XML"
   ClientHeight    =   7830
   ClientLeft      =   3750
   ClientTop       =   1935
   ClientWidth     =   15120
   ForeColor       =   &H00505050&
   Icon            =   "frmxml.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7830
   ScaleWidth      =   15120
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "\"
      Height          =   945
      Left            =   135
      TabIndex        =   9
      Top             =   165
      Width           =   10035
      Begin VB.TextBox txtDataHoraFiltro 
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
         Left            =   4290
         MaxLength       =   100
         TabIndex        =   14
         ToolTipText     =   "Descrição Menu"
         Top             =   450
         Width           =   1740
      End
      Begin VB.TextBox txtNumeroNota 
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
         Left            =   1275
         MaxLength       =   100
         TabIndex        =   11
         ToolTipText     =   "Descrição Menu"
         Top             =   450
         Width           =   1740
      End
      Begin CentroDeDistribuicao.chameleonButton cmdFiltrar 
         Height          =   315
         Left            =   6315
         TabIndex        =   13
         Top             =   450
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         BTYPE           =   14
         TX              =   "Filtrar"
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
         MICON           =   "frmxml.frx":23FA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Data e Hora"
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   3330
         TabIndex        =   15
         Top             =   510
         Width           =   2805
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Número de NF"
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   150
         TabIndex        =   12
         Top             =   510
         Width           =   2805
      End
      Begin VB.Label lblFiltro 
         BackStyle       =   0  'Transparent
         Caption         =   "Filtro de LOG XML"
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
         Height          =   285
         Left            =   150
         TabIndex        =   10
         Top             =   150
         Width           =   4470
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3480
      Top             =   7170
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "\"
      Height          =   6480
      Left            =   10290
      TabIndex        =   7
      Top             =   165
      Width           =   4740
      Begin VB.Label lblLogTempo 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         ForeColor       =   &H00E0E0E0&
         Height          =   615
         Left            =   150
         TabIndex        =   8
         Top             =   150
         Width           =   4470
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   44
      Left            =   150
      ScaleHeight     =   45
      ScaleWidth      =   14880
      TabIndex        =   3
      Top             =   6810
      Width           =   14880
   End
   Begin ComctlLib.ProgressBar ProgressBarPrincipal 
      Height          =   255
      Left            =   495
      TabIndex        =   0
      Top             =   6390
      Width           =   9660
      _ExtentX        =   17039
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.TextBox txtStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00505050&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   4.5
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   130
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   6390
      Width           =   210
   End
   Begin MSFlexGridLib.MSFlexGrid grdLog 
      Height          =   4980
      Left            =   135
      TabIndex        =   1
      Top             =   1275
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   8784
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   5263440
      ForeColor       =   16777215
      BackColorFixed  =   -2147483635
      ForeColorFixed  =   16777215
      ForeColorSel    =   5263440
      BackColorBkg    =   5263440
      GridColor       =   5263440
      GridColorFixed  =   5263440
      ScrollBars      =   2
      AllowUserResizing=   2
      Appearance      =   0
   End
   Begin VB.Timer timerProcurarXML 
      Enabled         =   0   'False
      Left            =   15060
      Top             =   3600
   End
   Begin CentroDeDistribuicao.chameleonButton cmdRetornar 
      Height          =   510
      Left            =   13620
      TabIndex        =   4
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
      MICON           =   "frmxml.frx":2416
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdBuscarNF 
      Height          =   510
      Left            =   10740
      TabIndex        =   5
      Top             =   6945
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Buscar XML"
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
      MICON           =   "frmxml.frx":2432
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CentroDeDistribuicao.chameleonButton cmdAtualizar 
      Height          =   510
      Left            =   12180
      TabIndex        =   6
      Top             =   6945
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Atualizar"
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
      MICON           =   "frmxml.frx":244E
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
Attribute VB_Name = "frmXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Editado por: Felip3FL (felip3.fl@gmail.com)
'Última atualização: 15/05/2013
'Versão: 1.1.31

Option Explicit

Private informacaoXML As String
Private nomeCamposXML(4) As String
'PRIAVTE posicao As Byte



Public arquivo As String


Dim sql As String
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset

Dim tamanhoCampoData As Integer

Private Sub chameleonButton1_Click()

End Sub

Private Sub cmdAtualizar_Click()
    carregaGridLog "", ""
    timerProcurarXML_Timer
    timerProcurarXML.Enabled = buscaAutomaticaXML
End Sub

Private Sub cmdFiltrar_Click()
    carregaGridLog txtNumeroNota.Text, txtDataHoraFiltro
End Sub

Private Sub cmdRetornar_Click()

    Unload Me
    
    If GLB_USU_Codigo = Empty Then
        sairDoSistema
    End If
    
End Sub

Private Sub Form_Activate()
    
    timerProcurarXML.Interval = Tempo
    timerProcurarXML.Enabled = buscaAutomaticaXML
    'timerProcurarXML_Timer
End Sub

Private Sub Form_Load()

    Dim arquivo As String

    carregarPosicaoTamanhoTela Me

    tamanhoCampoData = 24
''\\128.0.22.27\d$\Alfasig\ArquivosNFe\60872124001322\arq1_20131022\-nfeent
    'pastaRecebido = "D:\Alfasig\ArquivosNFe\60872124001322\arq1_20131022\-nfeent\"
    'pastaLido = "D:\Alfasig\ArquivosNFe\60872124001322\arq1_20131022\-nfeent\Lido\"
    'pastaInvalido = "D:\Alfasig\ArquivosNFe\60872124001322\arq1_20131022\-nfeent\Invalido\"

    'pastaRecebido = "\\128.0.22.27\d$\Alfasig\ArquivosNFe\60872124001322\arq1_20131022\-nfeent\"
    'pastaLido = "\\128.0.22.27\d$\Alfasig\ArquivosNFe\60872124001322\arq1_20131022\-nfeent\Lido\"
    'pastaInvalido = "\\128.0.22.27\d$\Alfasig\ArquivosNFe\60872124001322\arq1_20131022\-nfeent\Invalido\"

    'CarregarDBIni
    grdLog.Rows = 0
    grdLog.ColWidth(0) = grdLog.Width
    
    lblLogTempo.Caption = "LOG Interno" & vbNewLine & vbNewLine
    
    If buscaAutomaticaXML Then
        lblLogTempo.Caption = lblLogTempo.Caption & " - Verificação a cada " & Mid(Tempo, 1, Len(Tempo) - 3) & " segundo(s)"
    Else
        lblLogTempo.Caption = lblLogTempo.Caption & " - Verificação automatica desativada"
    End If
    
    carregaGridLog "", ""
    
End Sub

Private Function NFeDemeo(ByRef fornecedor As String) As Boolean
    If fornecedor = "DEMEO" Then
        NFeDemeo = True
    Else
        NFeDemeo = False
    End If
End Function

Private Sub cmdGuardaDadosBanco_Click()
    
    Dim valido As Boolean
    
    Screen.MousePointer = 11
    ProgressBarPrincipal.Value = 0
    valido = False
    
    deletaDadosTabelas
    
    If Not NFeDemeo(nomeCamposXML(4)) Then
        valido = True
    Else
        valido = False
        mensagemLOG 2, "NFe " & nomeCamposXML(1) & " com CNPJ da De Meo. Arquivo deletado."
        Kill pastaRecebido & arquivo
    End If
    
    If Not NFeRepetido(nomeCamposXML(1), nomeCamposXML(2), nomeCamposXML(4)) And valido = True Then
        
        carregarDadosUsuario
        
        ProgressBarPrincipal.Value = 25
    
        If carregarDadosNFe_Ide And carregarDadosNFe_emit And carregarDadosNFe_dest And _
        carregarDadosNFe_prod And carregarDadosNFe_total And carregarDadosNFe_transp And _
        carregarDadosNFe_cobr And carregarDadosNFe_infAdic Then
            
            mensagemLOG 1, transferirDadosTabelas(nomeCamposXML(0), nomeCamposXML(1), nomeCamposXML(2), nomeCamposXML(4))
            trocaArquivoPasta2 pastaRecebido, pastaLido
                        
            'Exit Sub
        Else
            mensagemLOG 3, "Erro na gravação da tabela"
            deletaDadosTabelas
            trocaArquivoPasta2 pastaRecebido, pastaInvalido
        End If

            ProgressBarPrincipal.Value = 75
    ElseIf valido = True Then
    
        If nomeCamposXML(1) = "" Then
            mensagemLOG 3, "Não foi possivel encontrar o número da Nota Fiscal (" & arquivo & ")"
        ElseIf nomeCamposXML(2) = "" Then
            mensagemLOG 3, "NFe " & nomeCamposXML(1) & " não pode ser armazenado. Serie não encontrado"
        ElseIf nomeCamposXML(4) = "" Then
            mensagemLOG 3, "NFe " & nomeCamposXML(1) & " não pode ser armazenado. Fornecedor não encontrado"
        Else
            mensagemLOG 3, "NFe " & nomeCamposXML(1) & " não pode ser armazenado."
        End If
        trocaArquivoPasta2 pastaRecebido, pastaInvalido
    End If
    
    ProgressBarPrincipal.Value = 100
    Screen.MousePointer = 0
    
End Sub


Private Sub ProcurarXML()

    informacaoXML = carregarArquivo
    
    If informacaoXML = "" Then
        timerProcurarXML.Enabled = True
    Else
        timerProcurarXML.Enabled = False
        
        If validaXML(informacaoXML) Then
            carregarDadosUsuario
            cmdGuardaDadosBanco_Click
        Else
            mensagemLOG 3, "XML invalido (" & arquivo & ")"
            trocaArquivoPasta2 pastaRecebido, pastaInvalido
        End If
        
        ProcurarXML
        
    End If

End Sub


Private Sub carregarDadosUsuario()
    
    nomeCamposXML(0) = Loja                                                             'Loja
    nomeCamposXML(1) = adquirirCampo("nNF", informacaoXML)                              'NFe
    nomeCamposXML(2) = "NE"                                                             'Serie
    nomeCamposXML(3) = "A"                                                              'Situacao
    nomeCamposXML(4) = codigoFornecedor(adquirirCampo("CNPJ", informacaoXML))           'Fornecedor
    
End Sub


Private Sub timerProcurarXML_Timer()

    ProcurarXML
    
    If txtStatus.BackColor = &H80000004 Then
        txtStatus.BackColor = &H8000000D
    Else
        txtStatus.BackColor = &H80000004
    End If
    ProgressBarPrincipal.Value = 0
    
End Sub

Public Function mensagemLOG(tipoStatus As Byte, Mensagem As String)
                   
    Dim status As String
    Dim corLinha As ColorConstants
                   
    Select Case tipoStatus
        Case 1
            status = "LOG"
            corLinha = vbBlack
        Case 2
            status = "AVISO"
            corLinha = &H808080
        Case 3
            status = "ERRO"
            corLinha = vbRed
    End Select
                   
    grdLog.AddItem "[" & CStr(time) & "] [" & status & "] " & Mensagem
    
    grdLog.row = grdLog.Rows - 1
    grdLog.TopRow = grdLog.row
    
    sql = "insert into logxml(lxml_dataHora,lxml_mensagem,lxml_tipoMensagem) " & _
          "values(getdate(),'" & Mensagem & "','" & tipoStatus & "')"
          
    ADO_Cn_CDLocal.Execute (sql)
    
    grdLog.Refresh
                   
End Function



Private Sub carregaGridLog(ByRef filtroNF As String, ByRef dataFiltro As String)

    Dim status As String
    Dim corLinha As ColorConstants

    Screen.MousePointer = 11
    
    sql = "delete logxml where lxml_datahora <= dateadd(DAY, -15, getdate())"
    ADO_Cn_CDLocal.Execute sql
    
    sql = "select top 500 lxml_dataHora as data," & _
          "lxml_mensagem as mensagem," & _
          "lxml_tipoMensagem as tipo " & _
          "from logXML " & _
          "where lxml_mensagem like '%" & filtroNF & "%'" & _
          "and lxml_dataHora like '%" & dataFiltro & "%'" & _
          "order by lxml_dataHora"
    
    rs.CursorLocation = adUseClient
    rs.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
        grdLog.Rows = 0
    
    With rs
    
        Do While Not rs.EOF
            Select Case rs("tipo")
                Case 1
                    status = "LOG"
                    corLinha = vbBlack
                Case 2
                    status = "AVISO"
                    corLinha = &H808080
                Case 3
                    status = "ERRO"
                    corLinha = vbRed
            End Select
        
            grdLog.AddItem "[" & rs("data") & "] [" & status & "] " & rs("mensagem")
            
            grdLog.row = grdLog.Rows - 1
            grdLog.TopRow = grdLog.row
            grdLog.CellForeColor = corLinha
                
            .MoveNext
        Loop
            
    End With
    
    rs.Close
    
    Screen.MousePointer = 0
    
End Sub


'###################################################################################################################
'##################################                                 ################################################
'###################################################################################################################



Private Function carregarDadosNFe_Ide() As Boolean
    
    Dim campoXML(21) As String
    Dim posicao As Byte
    Dim tipoData As Boolean
    
    campoXML(0) = "cUF":        campoXML(1) = "cNF":        campoXML(2) = "natOp":
    campoXML(3) = "indPag":     campoXML(4) = "mod":        campoXML(5) = "serie":
    campoXML(6) = "nNF":        campoXML(7) = "dhEmi":      campoXML(8) = "dhSaiEnt":
    campoXML(9) = "hSaiEnt":    campoXML(10) = "tpNF":      campoXML(11) = "cMunFG":
    campoXML(12) = "tpImp":     campoXML(13) = "tpEmis":    campoXML(14) = "cDV":
    campoXML(15) = "tpAmb":     campoXML(16) = "finNFe":    campoXML(17) = "procEmi":
    campoXML(18) = "verProc":   campoXML(19) = "dhCont":    campoXML(20) = "xJust":
    campoXML(21) = "chNFe"
    
    For posicao = 0 To UBound(campoXML)
    
        If Mid(campoXML(posicao), 1, 1) = "d" Then
            tipoData = True
        Else
            tipoData = False
        End If
        
        campoXML(posicao) = adquirirCampo(campoXML(posicao), informacaoXML)
        
        If tipoData = True Then campoXML(posicao) = Mid(campoXML(posicao), 1, 10)
        
    Next posicao
    
    carregarDadosNFe_Ide = GuardaDadosBanco(nomeCamposXML, campoXML, "ide")
    informacaoXML = eliminarCamposLido("ide", informacaoXML)

End Function



Private Function carregarDadosNFe_emit() As Boolean
   
    Dim campoXML(18) As String
    Dim posicao As Byte
    
    campoXML(0) = "CNPJ":       campoXML(10) = "CEP":
    campoXML(1) = "xNome":      campoXML(11) = "cPais":
    campoXML(2) = "xFant":      campoXML(12) = "xPais":
    campoXML(3) = "xLgr":       campoXML(13) = "fone":
    campoXML(4) = "nro":        campoXML(14) = "IE":
    campoXML(5) = "xCpl":       campoXML(15) = "IEST":
    campoXML(6) = "xBairro":    campoXML(16) = "IM":
    campoXML(7) = "cMun":       campoXML(17) = "CNAE":
    campoXML(8) = "xMun":       campoXML(18) = "CRT":
    campoXML(9) = "UF"

    For posicao = 0 To UBound(campoXML)
        campoXML(posicao) = adquirirCampo(campoXML(posicao), informacaoXML)
    Next posicao
    
    carregarDadosNFe_emit = GuardaDadosBanco(nomeCamposXML, campoXML, "emit")
    informacaoXML = eliminarCamposLido("emit", informacaoXML)

End Function



Private Function carregarDadosNFe_dest() As Boolean
   
    Dim campoXML(16) As String
    Dim posicao As Byte
    
    campoXML(0) = "CNPJ":       campoXML(11) = "cPais":
    campoXML(1) = "CPF":        campoXML(12) = "xPais":
    campoXML(2) = "xNome":      campoXML(13) = "fone":
    campoXML(3) = "xLgr":       campoXML(14) = "IE":
    campoXML(4) = "nro":        campoXML(15) = "ISUF":
    campoXML(5) = "xCpl":       campoXML(16) = "email":
    campoXML(6) = "xBairro":
    campoXML(7) = "cMun":
    campoXML(8) = "xMun":
    campoXML(9) = "UF":
    campoXML(10) = "CEP":
    
    For posicao = 0 To UBound(campoXML)
        campoXML(posicao) = adquirirCampo(campoXML(posicao), informacaoXML)
    Next posicao
    
    carregarDadosNFe_dest = GuardaDadosBanco(nomeCamposXML, campoXML, "dest")
    informacaoXML = eliminarCamposLido("dest", informacaoXML)

End Function



Private Function carregarDadosNFe_prod() As Boolean

    Dim inicioCampo As Integer
    Dim campoXML(77) As String
    Dim posicao As Byte
    Dim informacaoXMLProduto As String
    
    Do While informacaoXML Like "*<det*"
    
        If informacaoXML Like "*<det nItem=*" Then
            inicioCampo = (InStr(informacaoXML, "<det nItem=")) + 12
        ElseIf informacaoXML Like "*<detnItem=*" Then
            inicioCampo = (InStr(informacaoXML, "<detnItem=")) + 11
        End If
        
        campoXML(0) = Mid$(informacaoXML, inicioCampo, (InStr(informacaoXML, Chr$(34) & ">")) - inicioCampo)
        
        If campoXML(0) = 23 Then
            Debug.Print "OK"
        End If
        
        informacaoXMLProduto = Mid$(informacaoXML, 1, (InStr(informacaoXML, "</det>")) + 5)
        
        ' ==================================================================================================================
        
        campoXML(1) = "cProd":       campoXML(2) = "cEAN":          campoXML(3) = "xProd":
        campoXML(4) = "NCM":         campoXML(5) = "EXTIPI":        campoXML(6) = "CFOP":
        campoXML(7) = "uCom":        campoXML(8) = "qCom":          campoXML(9) = "vUnCom":
        campoXML(10) = "vProd":      campoXML(11) = "cEANTrib":     campoXML(12) = "uTrib":
        campoXML(13) = "qTrib":      campoXML(14) = "vUnTrib":      campoXML(15) = "vFrete":
        campoXML(16) = "vSeg":       campoXML(17) = "vDesc":        campoXML(18) = "vOutro":
        campoXML(19) = "indTot"
        
        For posicao = 1 To 19
                campoXML(posicao) = adquirirCampo(campoXML(posicao), informacaoXMLProduto)
        Next posicao
        
        ' ==================================================================================================================
        
        campoXML(20) = "orig":       campoXML(21) = "CST":          campoXML(22) = "modBC":
        campoXML(23) = "vBC":        campoXML(24) = "pRedBC":       campoXML(25) = "pICMS":
        campoXML(26) = "vICMS":      campoXML(27) = "modBC":        campoXML(28) = "pMVAST":
        campoXML(29) = "pRedBCST":     campoXML(30) = "vBCST":        campoXML(31) = "pICMSST":
        campoXML(32) = "vICMSST"
        
        For posicao = 20 To 32
                campoXML(posicao) = adquirirCampo(campoXML(posicao), adquirirCampoImposto("ICMS", informacaoXMLProduto))
        Next posicao

        ' ==================================================================================================================
        
        campoXML(33) = "cIEnq":    campoXML(34) = "CNPJProd":      campoXML(35) = "cSelo":
        campoXML(36) = "qSelo":    campoXML(37) = "cEnq":          campoXML(38) = "CST":
        campoXML(39) = "vBC":      campoXML(40) = "qUnid":         campoXML(41) = "vUnid":
        campoXML(42) = "pIPI":     campoXML(43) = "vIPI":          campoXML(44) = "CST":
        
        For posicao = 33 To 44
            campoXML(posicao) = adquirirCampo(campoXML(posicao), adquirirCampoImposto("IPI", informacaoXMLProduto))
        Next posicao
        
        ' ==================================================================================================================
        
        campoXML(45) = "vBC":     campoXML(46) = "vDespAdu":      campoXML(47) = "vII":
        campoXML(48) = "vIOF"
        
        For posicao = 45 To 48
            campoXML(posicao) = adquirirCampo(campoXML(posicao), adquirirCampoImposto("II", informacaoXMLProduto))
        Next posicao
        
        ' ==================================================================================================================
        
        campoXML(49) = "CST": campoXML(50) = "vBC": campoXML(51) = "pPIS":
        campoXML(52) = "qBCProd": campoXML(53) = "vAliqProd": campoXML(54) = "vPIS":
        
        For posicao = 49 To 54
            campoXML(posicao) = adquirirCampo(campoXML(posicao), adquirirCampoImposto("PIS", informacaoXMLProduto))
        Next posicao
        
        ' ==================================================================================================================
        
        campoXML(55) = "vBC": campoXML(56) = "pPIS": campoXML(57) = "qBCProd":
        campoXML(58) = "vAliqProd": campoXML(59) = "vPIS"
        
        For posicao = 55 To 59
                campoXML(posicao) = adquirirCampo(campoXML(posicao), adquirirCampoImposto("PISST", informacaoXMLProduto))
        Next posicao
        
        ' ==================================================================================================================
        
        campoXML(60) = "CST":        campoXML(61) = "vBC":           campoXML(62) = "pCOFINS":
        campoXML(63) = "qBCProd":    campoXML(64) = "vAliqProd":     campoXML(65) = "vCOFINS"
        
        For posicao = 60 To 65
            campoXML(posicao) = adquirirCampo(campoXML(posicao), adquirirCampoImposto("COFINS", informacaoXMLProduto))
        Next posicao
        
        ' ==================================================================================================================
        
        campoXML(66) = "vBC":          campoXML(67) = "pCOFINS":         campoXML(68) = "qBCProd":
        campoXML(69) = "vAliqProd":    campoXML(70) = "vCOFINS"
        
        For posicao = 66 To 70
            campoXML(posicao) = adquirirCampo(campoXML(posicao), adquirirCampoImposto("COFINSST", informacaoXMLProduto))
        Next posicao
        
        ' ==================================================================================================================
        
        campoXML(71) = "vBC":          campoXML(72) = "vAliq":         campoXML(73) = "vISSQN":
        campoXML(74) = "cMunFG":       campoXML(75) = "cListServ":     campoXML(76) = "cListServ"
        
        For posicao = 71 To 76
            campoXML(posicao) = adquirirCampo(campoXML(posicao), adquirirCampoImposto("ISSQN", informacaoXMLProduto))
        Next posicao
        
        ' ==================================================================================================================
        
        campoXML(77) = adquirirCampo("infAdProd", informacaoXMLProduto)
        
        
        ' ==================================================================================================================
                
        
        carregarDadosNFe_prod = GuardaDadosBanco(nomeCamposXML, campoXML, "prod")
        informacaoXML = eliminarCamposLido("det", informacaoXML)

    Loop

End Function



Private Function carregarDadosNFe_total() As Boolean
   
   Dim campoXML(24) As String
   Dim posicao As Byte
   
   campoXML(0) = "vBC":         campoXML(1) = "vICMS":          campoXML(2) = "vBCST":
   campoXML(3) = "vST":         campoXML(4) = "vProd":          campoXML(5) = "vFrete":
   campoXML(6) = "vSeg":        campoXML(7) = "vDesc":          campoXML(8) = "vII":
   campoXML(9) = "vIPI":        campoXML(10) = "vCOFINS":       campoXML(11) = "vOutro":
   campoXML(12) = "vNF":        campoXML(13) = "vServ":         campoXML(14) = "vBCISSQ":
   campoXML(15) = "vISS":       campoXML(16) = "vPIS":          campoXML(17) = "vCOFINsISSQ":
   campoXML(18) = "vRetPIS":    campoXML(19) = "vRetCOFINS":    campoXML(20) = "vRetCSLL":
   campoXML(21) = "vBCIRRF":    campoXML(22) = "vIRRF":         campoXML(23) = "vBCIRRF":
   campoXML(24) = "vBCIRRF"
    
    For posicao = 0 To UBound(campoXML)
        campoXML(posicao) = adquirirCampo(campoXML(posicao), informacaoXML)
    Next posicao
    
    carregarDadosNFe_total = GuardaDadosBanco(nomeCamposXML, campoXML, "total")
    informacaoXML = eliminarCamposLido("total", informacaoXML)

End Function



Private Function carregarDadosNFe_transp() As Boolean
   
   Dim campoXML(24) As String
   Dim posicao As Byte
   
   campoXML(0) = "modFrete":        campoXML(1) = "CNPJ":           campoXML(2) = "CPF":
   campoXML(3) = "xNome":           campoXML(4) = "IE":             campoXML(5) = "xEnder":
   campoXML(6) = "xMun":            campoXML(7) = "UF":             campoXML(8) = "vServ":
   campoXML(9) = "vBCRet":          campoXML(10) = "pICMSRet":      campoXML(11) = "vICMSRet":
   campoXML(12) = "CFOP":           campoXML(13) = "cMunFG":        campoXML(14) = "placa":
   campoXML(15) = "UFveic":         campoXML(16) = "RNTC":          campoXML(17) = "qVol":
   campoXML(18) = "esp":            campoXML(19) = "marca":         campoXML(20) = "nVol":
   campoXML(21) = "pesoL":          campoXML(22) = "pesoB":         campoXML(23) = "lacres":
   campoXML(24) = "nLacres"
   
    Dim NomeCampo As String * 6
    NomeCampo = "transp"
    
    For posicao = 0 To UBound(campoXML)
        campoXML(posicao) = adquirirCampo(campoXML(posicao), informacaoXML)
    Next posicao
    
    carregarDadosNFe_transp = GuardaDadosBanco(nomeCamposXML, campoXML, NomeCampo)
    informacaoXML = eliminarCamposLido(NomeCampo, informacaoXML)
    
End Function



Private Function carregarDadosNFe_cobr() As Boolean
    
    Dim campoXML(6) As String
    Dim posicao As Byte
    
    Dim NomeCampo As String * 4
    NomeCampo = "cobr"
    
    If verificarCampoExiste(informacaoXML, NomeCampo) Then
    
        Do While informacaoXML Like "*<dup>*"

            campoXML(0) = "nFat":        campoXML(1) = "vOrig":           campoXML(2) = "vDesc":
            campoXML(3) = "vLiq":           campoXML(4) = "nDup":             campoXML(5) = "dVenc":
            campoXML(6) = "vDup"
        
            For posicao = 0 To 3
                campoXML(posicao) = adquirirCampo(campoXML(posicao), informacaoXML)
            Next posicao
            
            For posicao = 4 To 6
                campoXML(posicao) = adquirirCampo(campoXML(posicao), adquirirCampoImposto("dup", informacaoXML))
            Next
            
            carregarDadosNFe_cobr = GuardaDadosBanco(nomeCamposXML, campoXML, NomeCampo)
            informacaoXML = eliminarCamposLidoEspecifico("dup", informacaoXML)
        
        Loop

    End If
    
    carregarDadosNFe_cobr = True

End Function



Private Function carregarDadosNFe_infAdic() As Boolean
   
       Dim campoXML(7) As String
       Dim posicao As Byte
       
       campoXML(0) = "infAdFisco":        campoXML(1) = "infCpl":           campoXML(2) = "xCampoCont":
       campoXML(3) = "xTextoCont":        campoXML(4) = "xCampoFisco":      campoXML(5) = "xTextoFisco":
       campoXML(6) = "nProc":             campoXML(7) = "indProc"
    
       For posicao = 0 To UBound(campoXML)
           campoXML(posicao) = adquirirCampo(campoXML(posicao), informacaoXML)
       Next posicao
        
       carregarDadosNFe_infAdic = GuardaDadosBanco(nomeCamposXML, campoXML, "infAdic")
       informacaoXML = eliminarCamposLido("infAdic", informacaoXML)

End Function


Private Sub verificarPastaExiste()
    
On Error GoTo TrataErro
    
    If Dir(pastaLido, vbDirectory) = "" Then
        MkDir pastaLido
    End If
    
    If Dir(pastaRecebido, vbDirectory) = "" Then
        MkDir pastaRecebido
    End If
    
TrataErro:
    Select Case Err.Number
        Case 76
        MsgBox "Não foi possível localizar ou criar as pastas:" & vbNewLine _
        & pastaLido & vbNewLine & pastaRecebido, vbCritical, "Erro"
        End
    End Select
    
End Sub

Public Function carregarArquivoBusca(endereco As String) As String
    
    Dim fso As New FileSystemObject
    Dim arquivoXML As TextStream
    
    arquivo = endereco
    If arquivo Like "*.xml" Or arquivo Like "*.XML" Then
    
            On Error GoTo erroAbrirXML
    
            Set arquivoXML = fso.OpenTextFile(pastaRecebido & arquivo)
            carregarArquivoBusca = arquivoXML.ReadAll
            carregarArquivoBusca = Replace(carregarArquivo, Chr(10), "")
            arquivoXML.Close
            
    Else
        carregarArquivoBusca = ""
    End If
    
erroAbrirXML:
    Select Case Err.Number
    Case 62
            carregarArquivoBusca = ""
            trocaArquivoPasta2 pastaRecebido, pastaInvalido
            
    End Select
    
End Function

Public Function carregarArquivo() As String
    
    Dim fso As New FileSystemObject
    Dim arquivoXML As TextStream
    
    arquivo = Dir(pastaRecebido & "*.xml")
    If arquivo Like "*.xml" Or arquivo Like "*.XML" Then
    
            On Error GoTo erroAbrirXML
    
            Set arquivoXML = fso.OpenTextFile(pastaRecebido & arquivo)
            carregarArquivo = arquivoXML.ReadAll
            carregarArquivo = Replace(carregarArquivo, Chr(10), "")
            arquivoXML.Close
            
    Else
        carregarArquivo = ""
    End If
    
erroAbrirXML:
    Select Case Err.Number
    Case 62
            carregarArquivo = ""
            trocaArquivoPasta2 pastaRecebido, pastaInvalido
            
    End Select
    
End Function



Public Function validaXML(informacaoXML As String) As Boolean

    Dim camposParaValidar(6) As String
    camposParaValidar(0) = "<NFe xmlns=" & Chr(34) & "http://www.portalfiscal.inf.br/nfe" & Chr(34) & "><nfeProc versao=" & Chr(34) & "2.00" & Chr(34) & " xmlns=" & Chr(34) & "http://www.portalfiscal.inf.br/nfe" & Chr(34) & ">"
    camposParaValidar(1) = "<NFe xmlns=" & Chr(34) & "http://www.portalfiscal.inf.br/nfe" & Chr(34) & "><nfeProc xmlns=" & Chr(34) & "http://www.portalfiscal.inf.br/nfe" & Chr(34) & " versao=" & Chr(34) & "2.00" & Chr(34) & ">"
    camposParaValidar(2) = "<NFe xmlns=" & Chr(34) & "http://www.portalfiscal.inf.br/nfe" & Chr(34) & "><protNFe versao=" & Chr(34) & "2.00" & Chr(34) & " xmlns=" & Chr(34) & "http://www.portalfiscal.inf.br/nfe" & Chr(34) & ">"
    camposParaValidar(3) = "<NFe xmlns=" & Chr(34) & "http://www.portalfiscal.inf.br/nfe" & Chr(34) & "><?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & "?>"
    camposParaValidar(4) = "<NFe xmlns=" & Chr(34) & "http://www.portalfiscal.inf.br/nfe" & Chr(34) & "><nfeProc versao=" & Chr(34) & "3.10" & Chr(34) & " xmlns=" & Chr(34) & "http://www.portalfiscal.inf.br/nfe" & Chr(34) & ">"
    camposParaValidar(5) = "<NFe xmlns=" & Chr(34) & "http://www.portalfiscal.inf.br/nfe" & Chr(34) & ">"
    camposParaValidar(6) = "http://www.portalfiscal.inf.br/nfe"

    Dim i As Integer
    For i = LBound(camposParaValidar) To UBound(camposParaValidar)
        If informacaoXML Like "*" & camposParaValidar(i) & "*" Then
            validaXML = True
            Exit Function
        Else
            validaXML = False
        End If
    Next
    
End Function


Public Function trocaArquivoPasta2(pastaOrigem As String, pastaDestino As String)

On Error GoTo TrataErro

    FileCopy pastaOrigem & arquivo, pastaDestino & arquivo
    Kill pastaOrigem & arquivo
    
TrataErro:
    Select Case Err.Number
        Case 53
        MsgBox "Arquivo XML lido não pode ser encontrado na pasta:" & vbNewLine & pastaOrigem & arquivo, _
        vbExclamation, "Arquivo não encontrado"
        Case 70
        MsgBox "Arquivo XML não pode ser apagado:" & vbNewLine & pastaOrigem & arquivo & _
        vbNewLine & "Remova manualmente o arquivo XML ou configure o UAC do Windows", vbCritical, "Erro de permissão"
        End
    End Select
    
End Function


Public Function trocaArquivoPasta()

On Error GoTo TrataErro

    FileCopy pastaRecebido & arquivo, pastaLido & arquivo
    Kill pastaRecebido & arquivo
    
TrataErro:
    Select Case Err.Number
        Case 53
        MsgBox "Arquivo XML lido não pode ser encontrado na pasta", vbExclamation, "Arquivo não encontrado"
    End Select
    
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''






Function GuardaDadosBanco(ByRef informacaoBasicaXML() As String, ByRef nomeCamposXML() As String, tabela As String) As Boolean
    Dim Data As String
    'ConectaADO
    'cn.Open

        Dim posicao As Byte
        'sql = "insert into NFe_entrada_" & tabela & " values ('" & nomeCamposXML(0).nome & "'"
        sql = "exec SP_Inserir_NFe_Entrada_" & tabela & " '" & informacaoBasicaXML(0) & "'"
        
        For posicao = 1 To UBound(informacaoBasicaXML)
            sql = sql + ", '" & informacaoBasicaXML(posicao) & "'"
        Next posicao
        
        For posicao = 0 To UBound(nomeCamposXML)
        If tabela = "ide" And posicao = 19 Then
        nomeCamposXML(posicao) = Mid(nomeCamposXML(posicao), 1, 10)
         sql = sql + ", '" & nomeCamposXML(posicao) & "'"
        Else
        If InStr(nomeCamposXML(posicao), "'") > 0 Then
        nomeCamposXML(posicao) = Replace(nomeCamposXML(posicao), "'", " ")
        End If
        
       
            sql = sql + ", '" & nomeCamposXML(posicao) & "'"
            End If
        Next posicao
        
        'sql = sql + ")"
        
        '-2147217833
        'On Error GoTo erroGravacaoSQL
        
        ADO_Cn_CDLocal.Execute (sql)
        GuardaDadosBanco = True
        
'erroGravacaoSQL:
    'cn.Close

    'Select Case Err.Number
    'Case -2147217833
            'GuardaDadosBanco = False
    'End Select
    
End Function


Function NFeRepetido(numeroNFe, serie, fornecedor As String) As Boolean
    
    If numeroNFe <> "" And serie <> "" And fornecedor <> "" Then
    
            sql = "select CC_NotaFiscal, CC_Serie, CC_Fornecedor from capaNFcompra where CC_NotaFiscal = " & numeroNFe
        
            'ConectaADO
            'cn.Open
            rs.CursorLocation = adUseClient
            rs.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
            With rs
                If .BOF And .EOF Then
                    NFeRepetido = False
                Else
                
                    Do While Not .EOF
                        If .Fields("CC_NotaFiscal") Like "*" & numeroNFe & "*" And _
                        .Fields("CC_Serie") Like "*" & serie & "*" And _
                        .Fields("CC_Fornecedor") Like "*" & fornecedor & "*" Then
                        
                            'mensagemLOG 2, "NFe " & numeroNFe & " já armazenado"
                            NFeRepetido = True
                            Exit Do
                            
                        Else
                            NFeRepetido = False
                        End If
                        .MoveNext
                    Loop
    
                End If
            End With
    
        rs.Close
        'cn.Close
    Else
        'lstLog.AddItem "[" & CStr(Time) & "] [AVISO] Erro de leitura do XML " & numeroNFe
        NFeRepetido = True
    End If

End Function

Function codigoFornecedor(cnpj As String) As String

    'ConectaAD'
    'cn.Open
    
        sql = "select FO_CodigoFornecedor from fornecedor where FO_CGC like '%" & cnpj & "' order by len(FO_CGC) "
    
        rs.CursorLocation = adUseClient
        rs.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
        With rs
            If .BOF And .EOF Then
                
                rs.Close
                
                sql = "select lo_cgc from loja where lo_cgc like '%" & cnpj & "'"
                rs.CursorLocation = adUseClient
                rs.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
                
                With rs
                    If .BOF And .EOF Then
                        codigoFornecedor = ""
                    Else
                        codigoFornecedor = "DEMEO"
                    End If
                End With
                
                rs.Close
                
            Else
                codigoFornecedor = .Fields("FO_CodigoFornecedor")
                
                rs.Close
                
            End If
        End With
        
    
    'cn.Close
    
End Function


Sub deletaDadosTabelas()

    'ConectaADO
    'cn.Open
        ADO_Cn_CDLocal.Execute ("exec SP_Deleta_NFe_Entrada")
    'cn.Close

End Sub

Function transferirDadosTabelas(Loja, nfe, serie, fornecedor As String) As String

    Dim sql As String

    'ConectaADO
    'cn.Open
        sql = "exec SP_CriaCapaNFCompra_Via_NFE_Entrada '" & Loja & "', " & nfe & ", '" & serie & "', '" & fornecedor & "'"
        ADO_Cn_CDLocal.Execute (sql)
        
        sql = "exec SP_CriaVencimentosFornecedor_Via_NFE_Entrada '" & nfe & "', '" & serie & "', '" & fornecedor & "'"
        ADO_Cn_CDLocal.Execute (sql)
    'cn.Close
    
    sql = "exec SP_ItemNFCompra_Duplicado '" & nomeCamposXML(1) & "','" & nomeCamposXML(2) & "','" & nomeCamposXML(4) & "'"
    ADO_Cn_CDLocal.Execute (sql)
    
    transferirDadosTabelas = "NFe " & nfe & " armazenada com sucesso"

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function eliminarCamposLidoEspecifico(nomeTabela, informacaoXML As String) As String

    Dim informacaoElimiar As String

         informacaoElimiar = Mid$(informacaoXML, 1, (InStr(informacaoXML, "<" & nomeTabela & ">")) - 1)
         eliminarCamposLidoEspecifico = informacaoElimiar & Mid$(informacaoXML, (InStr(informacaoXML, "</" & nomeTabela & ">") + Len(nomeTabela) + 3), Len(informacaoXML))

End Function

Function eliminarCamposLido(nomeTabela, informacaoXML As String) As String
            
     eliminarCamposLido = Mid$(informacaoXML, (InStr(informacaoXML, "</" & nomeTabela & ">") + Len(nomeTabela) + 3), Len(informacaoXML))

End Function


Function adquirirCampo(campoXML, informacaoXML As String) As String
           
        If informacaoXML Like "*<" & campoXML & ">*" Then
        
            Dim inicioCampo, fimCampo As Integer
    
            inicioCampo = (InStr(informacaoXML, "<" & campoXML & ">")) + (Len(campoXML)) + 2
            fimCampo = (InStr(inicioCampo, informacaoXML, "</" & campoXML & ">") - inicioCampo)
    
            If inicioCampo + fimCampo <> 0 Then
                adquirirCampo = Mid$(informacaoXML, inicioCampo, fimCampo)
            End If
            
        Else
    
            adquirirCampo = ""
        
        End If

End Function


Function adquirirCampoImposto(campoImpostoXML, informacaoXML As String) As String
           
        Dim informacaoImposto As String
           
        If informacaoXML Like "*<" & campoImpostoXML & ">*" Then
                 
            Dim inicioCampo, fimCampo As Integer
            
            inicioCampo = InStr(informacaoXML, "<" & campoImpostoXML & ">")
            fimCampo = (InStr(informacaoXML, "</" & campoImpostoXML & ">") - InStr(informacaoXML, "<" & campoImpostoXML & ">"))
            
            informacaoImposto = Mid$(informacaoXML, inicioCampo, fimCampo)
            
            adquirirCampoImposto = informacaoImposto
        
        Else
        
            adquirirCampoImposto = ""
        
        End If

End Function


Public Function Replace(Texto As String, caracter As String, caracterParaSubstituir As String) As String
    
    Do While Texto Like "*" & caracter & "*"
        Texto = left$(Texto, (InStr(Texto, caracter) - 1)) _
        & caracterParaSubstituir _
        & right$(Texto, (Len(Texto) - (InStr(Texto, caracter))))
    Loop
    
    Replace = Texto
    
End Function


Public Function verificarCampoExiste(informacaoXML As String, campo As String) As Boolean
    If informacaoXML Like "*<" & campo & ">*" Then
        verificarCampoExiste = True
    Else
        verificarCampoExiste = False
    End If
End Function

Private Sub cmdBuscarNF_Click()

    Dim arquivoXMLTexto As String
    Dim fso As New FileSystemObject
    Dim arquivoXML As TextStream


'ricardo
CommonDialog1.Filter = "Arquivos XML|*.XML"
CommonDialog1.InitDir = "C:\"
CommonDialog1.ShowOpen

arquivoXMLTexto = CommonDialog1.FileName


   On Error GoTo erroAbrirXML
    
            Set arquivoXML = fso.OpenTextFile(arquivoXMLTexto)
            arquivoXMLTexto = arquivoXML.ReadAll
            arquivoXMLTexto = Replace(arquivoXMLTexto, Chr(10), "")
            arquivoXML.Close
            
erroAbrirXML:
    Select Case Err.Number
    Case 62
            arquivoXMLTexto = ""
            trocaArquivoPasta2 pastaRecebido, pastaInvalido
            
    End Select
    informacaoXML = arquivoXMLTexto
    
    If informacaoXML = "" Then
        timerProcurarXML.Enabled = True
    Else
        timerProcurarXML.Enabled = False
        
        If validaXML(informacaoXML) Then
            carregarDadosUsuario
            cmdGuardaDadosBanco_Click
        Else
            mensagemLOG 3, "XML invalido (" & arquivo & ")"
            trocaArquivoPasta2 pastaRecebido, pastaInvalido
        End If

        ProcurarXML
        
    End If

End Sub



Private Sub txtDataHoraFiltro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdFiltrar_Click
End Sub

Private Sub txtNumeroNota_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdFiltrar_Click
End Sub
