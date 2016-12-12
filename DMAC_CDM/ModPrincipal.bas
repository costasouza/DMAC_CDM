Attribute VB_Name = "ModPrincipal"
Option Explicit

Global ConexaoDLLaDO As New DMACD.conexaoADO

'Global nomeImpressora As Object



Public posicaoTelaY As String
Public posicaoTelaX As String
Public tamanhoTelaY As String
Public tamanhoTelaX As String

Public Type resolucaoTela
    Linhas As Single
    Colunas As Single
End Type

Public resolucaoOriginal As resolucaoTela

Public Type InformacaoCampoXML
    NOME As String
    campoLido As String
End Type


Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1
Const SWP_WNDFLAGS = SWP_NOMOVE Or SWP_NOSIZE
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2



' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '

Public ResX As Single
Public ResY As Single
Public OldX As Single
Public OldY As Single
Public resolucao As Boolean

'muda data e símbolo de R$
Public Const LOCALE_SSHORTDATE = &H1F
Public Const LOCALE_SCURRENCY = 20
Public Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Public Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Boolean

' muda resolução do vídeo
Public Type RECT
   left As Long
   top As Long
   right As Long
   bottom As Long
End Type

Public Declare Function GetClipCursor Lib "user32.dll" (lprc As RECT) As Long

Private Declare Function EnumDisplaySettings Lib "user32" Alias _
"EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, _
lpDevMode As Any) As Boolean

Private Declare Function ChangeDisplaySettings Lib "user32" Alias _
"ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwflags As Long) As Long

Const CCDEVICENAME = 32
Const CCFORMNAME = 32
Const DM_PELSWIDTH = &H80000
Const DM_PELSHEIGHT = &H100000

Private Type DEVMODE
   dmDeviceName As String * CCDEVICENAME
   dmSpecVersion As Integer
   dmDriverVersion As Integer
   dmSize As Integer
   dmDriverExtra As Integer
   dmFields As Long
   dmOrientation As Integer
   dmPaperSize As Integer
   dmPaperLength As Integer
   dmPaperWidth As Integer
   dmScale As Integer
   dmCopies As Integer
   dmDefaultSource As Integer
   dmPrintQuality As Integer
   dmColor As Integer
   dmDuplex As Integer
   dmYResolution As Integer
   dmTTOption As Integer
   dmCollate As Integer
   dmFormName As String * CCFORMNAME
   dmUnusedPadding As Integer
   dmBitsPerPel As Integer
   dmPelsWidth As Long
   dmPelsHeight As Long
   dmDisplayFlags As Long
   dmDisplayFrequency As Long
End Type

Dim DevM As DEVMODE

Public Sub AlterarResolucao(iWidth As Single, iHeight As Single)

   If Glb_AlteraResolucao = True Then
        
       Dim a As Boolean
       Dim i As Long
       Do
          a = EnumDisplaySettings(0&, i, DevM)
          i = i + 1
       Loop Until (a = False)
    
       Dim B As Long
       DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
       DevM.dmPelsWidth = iWidth
       DevM.dmPelsHeight = iHeight
       B = ChangeDisplaySettings(DevM, 0)
   
   End If
   
End Sub

' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' ' '

Sub JanelaTOP(form_top As Form)
   Call SetWindowPos(form_top.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_WNDFLAGS)
   'res = SetWindowPos(form_top.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
End Sub

Sub JanelaNORMAL(form_normal As Form)
    Call SetWindowPos(form_normal.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_WNDFLAGS)
End Sub


Public Function resolucaoTelaIdeal() As Boolean
    If resolucaoTela.Colunas = "1024" And resolucaoTela.Linhas = "768" Then
        resolucaoTelaIdeal = True
    Else
        resolucaoTelaIdeal = False
    End If
End Function

Public Function resolucaoTela() As resolucaoTela
    resolucaoTela.Linhas = Screen.Height / Screen.TwipsPerPixelX
    resolucaoTela.Colunas = Screen.Width / Screen.TwipsPerPixelY
End Function

Sub limpaGrid(ByRef GradeUsu)
    
    GradeUsu.Rows = GradeUsu.FixedRows
    GradeUsu.AddItem ""
    GradeUsu.RemoveItem GradeUsu.FixedRows

End Sub

Public Function gridVazio(ByRef grid) As Boolean
    If grid.Rows = grid.FixedRows Then
        gridVazio = True
    Else
        gridVazio = False
    End If
End Function

Public Sub carregarPosicaoTamanhoTela(tela As Form)
   tela.top = posicaoTelaY
   tela.left = posicaoTelaX
   tela.Height = tamanhoTelaY
   tela.Width = tamanhoTelaX
End Sub

Public Sub carregarPosicaoFrame(tela As Frame)
   tela.top = (tamanhoTelaY / 2) - (tela.Height / 2)
   tela.left = (tamanhoTelaX / 2) - (tela.Width / 2)
End Sub

Public Sub carregarPosicaoTela(tela As Form)
   tela.top = posicaoTelaY + ((tamanhoTelaY / 2) - (tela.Height / 2))
   tela.left = (tamanhoTelaX / 2) - (tela.Width / 2)
End Sub

Public Sub sairDoSistema()
    Call AlterarResolucao(resolucaoOriginal.Colunas, resolucaoOriginal.Linhas)
    End
End Sub


Public Function endIMGBotao(nomeBotao As String) As String
    
    Dim arquivo As String
    Dim enderecoArquivo As String
    
    enderecoArquivo = "c:\sistemas\dmac cdm\data\" & nomeBotao & ""
    arquivo = Dir(enderecoArquivo, vbDirectory)
    
    If arquivo = Empty Then
        enderecoArquivo = "c:\sistemas\dmac cdm\data\" & "btDefalt"
    Else
        enderecoArquivo = "c:\sistemas\dmac cdm\data\" & nomeBotao & ""
    End If
    
    endIMGBotao = enderecoArquivo
    
End Function

Public Function endIMG(nomeBotao As String) As String
    
    Dim arquivo As String
    Dim enderecoArquivo As String
    
    If GLB_modoOffline Then nomeBotao = nomeBotao & "off"
    
    enderecoArquivo = "c:\sistemas\dmac cdm\imagens\lojas\" & GLB_logoPedido & "_" & nomeBotao
    arquivo = Dir(enderecoArquivo, vbDirectory)
    
    If arquivo = Empty Then
        enderecoArquivo = "c:\sistemas\dmac cdm\imagens\" & nomeBotao
    End If
    
    endIMG = enderecoArquivo
    
End Function

Public Sub verificaAppExecucao()
    If App.PrevInstance Then
       MsgBox App.EXEName + " Já está executando", vbCritical
       End
    End If
End Sub

Public Function cancelarNota(ByRef nf As String, ByRef serie As String) As Boolean
    Dim adoCancelamento As New ADODB.Recordset
    Dim sql As String
    
    cancelarNota = False
    
    sql = "select nf as nf " & vbNewLine & _
          "from nfcapa " & vbNewLine & _
          "where serie = '" & serie & "' and nf = '" & nf & "'"
          
    With adoCancelamento
        .CursorLocation = adUseClient
        .Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
        
        If Not adoCancelamento.EOF Then
            sql = "exec SP_Cancela_NotaFiscal '" & nf & "', '" & serie & "'"
            ADO_Cn_CDLocal.Execute sql
            cancelarNota = True
        Else
            MsgBox "Nenhuma nota encontrada!", vbExclamation, "Cancelamento"
        End If
        
        .Close
        
    End With
End Function
