VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmPrincipal 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DMAC - Leitura XML"
   ClientHeight    =   4635
   ClientLeft      =   9285
   ClientTop       =   2310
   ClientWidth     =   5070
   Icon            =   "frmPrincipal.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   5070
   Begin VB.TextBox txtStatus 
      BackColor       =   &H80000004&
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
      Top             =   4200
      Width           =   210
   End
   Begin MSFlexGridLib.MSFlexGrid grdLog 
      Height          =   3960
      Left            =   135
      TabIndex        =   1
      Top             =   105
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   6985
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   -2147483635
      ForeColorSel    =   16777215
      BackColorBkg    =   16777215
      GridColor       =   16777215
      GridColorFixed  =   16777215
      AllowUserResizing=   2
   End
   Begin ComctlLib.ProgressBar ProgressBarPrincipal 
      Height          =   255
      Left            =   435
      TabIndex        =   0
      Top             =   4200
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Timer timerProcurarXML 
      Left            =   4470
      Top             =   3600
   End
End
Attribute VB_Name = "frmPrincipal"
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


Private Sub Form_Load()
    CarregarDBIni
    grdLog.Rows = 0
    grdLog.ColWidth(0) = 4730
    mensagemLOG 2, "Verificação a cada " & Mid(tempo, 1, Len(tempo) - 3) & " segundo(s)"
    
    timerProcurarXML.Interval = tempo
    timerProcurarXML_Timer
End Sub


Private Sub cmdGuardaDadosBanco_Click()
    
    Screen.MousePointer = 11
    ProgressBarPrincipal.Value = 0
    
    deletaDadosTabelas
    
    If Not NFeRepetido(nomeCamposXML(1), nomeCamposXML(2), nomeCamposXML(4)) Then
    
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
    Else
        mensagemLOG 3, "NFe " & nomeCamposXML(1) & " não pode ser armazenado"
        trocaArquivoPasta2 pastaRecebido, pastaLido
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

Private Sub grdLog_DblClick()
    If MsgBox("Deseja limpar o Log?", vbQuestion + vbYesNo, "Limpar Log") = 6 Then
        grdLog.Rows = 0
    End If
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

Public Function mensagemLOG(tipoStatus As Byte, mensagem As String)

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
                   
    grdLog.AddItem "[" & CStr(Time) & "] [" & status & "] " & mensagem
    
    grdLog.Row = grdLog.Rows - 1
    grdLog.TopRow = grdLog.Row
    grdLog.CellForeColor = corLinha
                   
End Function



'###################################################################################################################
'##################################                                 ################################################
'###################################################################################################################



Private Function carregarDadosNFe_Ide() As Boolean
    
    Dim campoXML(20) As String
    Dim posicao As Byte
    
    campoXML(0) = "cUF":        campoXML(1) = "cNF":        campoXML(2) = "natOp":
    campoXML(3) = "indPag":     campoXML(4) = "mod":        campoXML(5) = "serie":
    campoXML(6) = "nNF":        campoXML(7) = "dEmi":       campoXML(8) = "dSaiEnt":
    campoXML(9) = "hSaiEnt":    campoXML(10) = "tpNF":      campoXML(11) = "cMunFG":
    campoXML(12) = "tpImp":     campoXML(13) = "tpEmis":    campoXML(14) = "cDV":
    campoXML(15) = "tpAmb":     campoXML(16) = "finNFe":    campoXML(17) = "procEmi":
    campoXML(18) = "verProc":   campoXML(19) = "dhCont":    campoXML(20) = "xJust"
    
    For posicao = 0 To UBound(campoXML)
        campoXML(posicao) = adquirirCampo(campoXML(posicao), informacaoXML)
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
    
    Do While informacaoXML Like "*<det*"
    
        If informacaoXML Like "*<det nItem=*" Then
            inicioCampo = (InStr(informacaoXML, "<det nItem=")) + 12
        ElseIf informacaoXML Like "*<detnItem=*" Then
            inicioCampo = (InStr(informacaoXML, "<detnItem=")) + 11
        End If
        
        campoXML(0) = Mid$(informacaoXML, inicioCampo, (InStr(informacaoXML, Chr$(34) & ">")) - inicioCampo)
        
        ' ==================================================================================================================
        
        campoXML(1) = "cProd":       campoXML(2) = "cEAN":          campoXML(3) = "xProd":
        campoXML(4) = "NCM":         campoXML(5) = "EXTIPI":        campoXML(6) = "CFOP":
        campoXML(7) = "uCom":        campoXML(8) = "qCom":          campoXML(9) = "vUnCom":
        campoXML(10) = "vProd":      campoXML(11) = "cEANTrib":     campoXML(12) = "uTrib":
        campoXML(13) = "qTrib":      campoXML(14) = "vUnTrib":      campoXML(15) = "vFrete":
        campoXML(16) = "vSeg":       campoXML(17) = "vDesc":        campoXML(18) = "vOutro":
        campoXML(19) = "indTot"
        
        For posicao = 1 To 19
                campoXML(posicao) = adquirirCampo(campoXML(posicao), informacaoXML)
        Next posicao
        
        ' ==================================================================================================================
        
        campoXML(20) = "orig":       campoXML(21) = "CST":          campoXML(22) = "modBC":
        campoXML(23) = "vBC":        campoXML(24) = "pRedBC":       campoXML(25) = "pICMS":
        campoXML(26) = "vICMS":      campoXML(27) = "modBC":        campoXML(28) = "pMVA":
        campoXML(29) = "pRedBC":     campoXML(30) = "vBC":          campoXML(31) = "pICMS":
        campoXML(32) = "vICMS"
        
        For posicao = 20 To 32
                campoXML(posicao) = adquirirCampo(campoXML(posicao), adquirirCampoImposto("ICMS", informacaoXML))
        Next posicao

        ' ==================================================================================================================
        
        campoXML(33) = "cIEnq":    campoXML(34) = "CNPJProd":      campoXML(35) = "cSelo":
        campoXML(36) = "qSelo":    campoXML(37) = "cEnq":          campoXML(38) = "CST":
        campoXML(39) = "vBC":      campoXML(40) = "qUnid":         campoXML(41) = "vUnid":
        campoXML(42) = "pIPI":     campoXML(43) = "vIPI":          campoXML(44) = "CST":
        
        For posicao = 33 To 44
            campoXML(posicao) = adquirirCampo(campoXML(posicao), adquirirCampoImposto("IPI", informacaoXML))
        Next posicao
        
        ' ==================================================================================================================
        
        campoXML(45) = "vBC":     campoXML(46) = "vDespAdu":      campoXML(47) = "vII":
        campoXML(48) = "vIOF"
        
        For posicao = 45 To 48
            campoXML(posicao) = adquirirCampo(campoXML(posicao), adquirirCampoImposto("II", informacaoXML))
        Next posicao
        
        ' ==================================================================================================================
        
        campoXML(49) = "CST": campoXML(50) = "vBC": campoXML(51) = "pPIS":
        campoXML(52) = "qBCProd": campoXML(53) = "vAliqProd": campoXML(54) = "vPIS":
        
        For posicao = 49 To 54
            campoXML(posicao) = adquirirCampo(campoXML(posicao), adquirirCampoImposto("PIS", informacaoXML))
        Next posicao
        
        ' ==================================================================================================================
        
        campoXML(55) = "vBC": campoXML(56) = "pPIS": campoXML(57) = "qBCProd":
        campoXML(58) = "vAliqProd": campoXML(59) = "vPIS"
        
        For posicao = 55 To 59
                campoXML(posicao) = adquirirCampo(campoXML(posicao), adquirirCampoImposto("PISST", informacaoXML))
        Next posicao
        
        ' ==================================================================================================================
        
        campoXML(60) = "CST":        campoXML(61) = "vBC":           campoXML(62) = "pCOFINS":
        campoXML(63) = "qBCProd":    campoXML(64) = "vAliqProd":     campoXML(65) = "vCOFINS"
        
        For posicao = 60 To 65
            campoXML(posicao) = adquirirCampo(campoXML(posicao), adquirirCampoImposto("COFINS", informacaoXML))
        Next posicao
        
        ' ==================================================================================================================
        
        campoXML(66) = "vBC":          campoXML(67) = "pCOFINS":         campoXML(68) = "qBCProd":
        campoXML(69) = "vAliqProd":    campoXML(70) = "vCOFINS"
        
        For posicao = 66 To 70
            campoXML(posicao) = adquirirCampo(campoXML(posicao), adquirirCampoImposto("COFINSST", informacaoXML))
        Next posicao
        
        ' ==================================================================================================================
        
        campoXML(71) = "vBC":          campoXML(72) = "vAliq":         campoXML(73) = "vISSQN":
        campoXML(74) = "cMunFG":       campoXML(75) = "cListServ":     campoXML(76) = "cListServ"
        
        For posicao = 71 To 76
            campoXML(posicao) = adquirirCampo(campoXML(posicao), adquirirCampoImposto("ISSQN", informacaoXML))
        Next posicao
        
        ' ==================================================================================================================
        campoXML(77) = adquirirCampo("infAdProd", informacaoXML)
        
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
   
    Dim nomeCampo As String * 6
    nomeCampo = "transp"
    
    For posicao = 0 To UBound(campoXML)
        campoXML(posicao) = adquirirCampo(campoXML(posicao), informacaoXML)
    Next posicao
    
    carregarDadosNFe_transp = GuardaDadosBanco(nomeCamposXML, campoXML, nomeCampo)
    informacaoXML = eliminarCamposLido(nomeCampo, informacaoXML)
    
End Function



Private Function carregarDadosNFe_cobr() As Boolean
    
    Dim campoXML(6) As String
    Dim posicao As Byte
    
    Dim nomeCampo As String * 4
    nomeCampo = "cobr"
    
    If verificarCampoExiste(informacaoXML, nomeCampo) Then
    
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
            
            carregarDadosNFe_cobr = GuardaDadosBanco(nomeCamposXML, campoXML, nomeCampo)
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
