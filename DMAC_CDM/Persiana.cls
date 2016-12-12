VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Persiana"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Type MSG
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Private Const PM_NOREMOVE = &H0

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WaitMessage Lib "user32" () As Long
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long

Private WithEvents pnlCopia As SSPanel
Attribute pnlCopia.VB_VarHelpID = -1

Public FundoNormal As Long
Public FundoDestacado As Long
Public LetraNormal As Long
Public LetraDestacada As Long
Public TamanhoMinimo As Long
Public TamanhoMaximo As Long
Public Step As Long
Public UsePonteiroMouse As Boolean

Public Event Destacou()
Public Event Normalizou()

Dim Formulario As Form
Dim Fator As Single
Dim PtIni As POINTAPI

Private mvarPonteiroMouse As Variant
Private mvarNomeFormulario As String
Private mvarNomePainel As String

Public Sub Destaca()

    If pnlCopia.BackColor <> FundoDestacado Then
        pnlCopia.BackColor = FundoDestacado
        pnlCopia.ForeColor = LetraDestacada
        RaiseEvent Destacou
    End If

End Sub

Public Sub Normaliza()
    
    If pnlCopia.BackColor <> FundoNormal Then
        pnlCopia.BackColor = FundoNormal
        pnlCopia.ForeColor = LetraNormal
        RaiseEvent Normalizou
    End If

End Sub

Public Property Let PonteiroMouse(ByVal vData As Variant)
    
    mvarPonteiroMouse = vData
    
    If UsePonteiroMouse Then
        pnlCopia.MousePointer = 99
        pnlCopia.MouseIcon = vData
    End If
    
End Property

Public Property Set PonteiroMouse(ByVal vData As Object)
    
    Set mvarPonteiroMouse = vData

End Property

Public Property Get PonteiroMouse() As Variant
    
    If IsObject(mvarPonteiroMouse) Then
        Set PonteiroMouse = mvarPonteiroMouse
    Else
        PonteiroMouse = mvarPonteiroMouse
    End If

End Property

Public Property Let NomePainel(ByVal vData As String)
    
    mvarNomePainel = vData
    Set pnlCopia = Formulario.Controls(vData)

End Property

Public Property Get NomePainel() As String
    
    NomePainel = mvarNomePainel

End Property

Public Property Let NomeFormulario(ByVal vData As String)
    
    Dim Indice As Long
    Dim Maximo As Long
    
    Maximo = Forms.Count - 1
    
    mvarNomeFormulario = ""
    
    For Indice = 0 To Maximo Step 1
        If UCase(Forms(Indice).Name) = UCase(vData) Then
            mvarNomeFormulario = vData
            Set Formulario = Forms(Indice)
            Exit For
        End If
    Next Indice
    
    If mvarNomeFormulario = "" Then
        Err.Raise -10, "Classe Cadastro", "Nome de formulario não encontrado."
    End If
    
End Property

Public Property Get NomeFormulario() As String
    
    NomeFormulario = mvarNomeFormulario

End Property

Private Sub Class_Initialize()
    
    Fator = Screen.TwipsPerPixelY

End Sub

Private Sub Mensagem()

    Dim Recebe As MSG
    Dim ret As Long
    Dim ret2 As Long
    Dim Mess As Long
    Dim NovoTamanho As Long
    
    ret = 0
        
    Do While ret = 0
        WaitMessage
        ret2 = PeekMessage(Recebe, Formulario.hwnd, 0, 0, PM_NOREMOVE)
        Mess = Recebe.message
        If Mess = &H200 Or Mess = &H118 Then 'MouseMove
            NovoTamanho = TamanhoMinimo + ((Recebe.pt.Y - PtIni.Y) * Fator)
            If NovoTamanho > TamanhoMinimo And NovoTamanho < TamanhoMaximo Then
                pnlCopia.Height = NovoTamanho
            End If
        ElseIf Mess = &H202 Then 'MouseUp botão Esquerdo
            Diminuir
            Exit Sub
        ElseIf Mess = &H204 Then 'MouseUp botão Direito
            Exit Sub
        End If
        Formulario.Refresh
        DoEvents
    Loop

End Sub

Private Sub Diminuir()

    Do While pnlCopia.Height > TamanhoMinimo
        pnlCopia.Height = pnlCopia.Height - Step
        Formulario.Refresh
    Loop
    pnlCopia.Height = TamanhoMinimo
    Normaliza
    Formulario.Refresh

End Sub

Private Sub pnlCopia_MouseDown(Button As Integer, shift As Integer, X As Single, Y As Single)
    
    If Button = vbLeftButton Then
        GetCursorPos PtIni
        Mensagem
    End If

End Sub

Private Sub pnlCopia_MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)
    
    Destaca

End Sub
