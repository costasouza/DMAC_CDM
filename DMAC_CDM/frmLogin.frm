VERSION 5.00
Begin VB.Form frmLogin 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Login CD"
   ClientHeight    =   10305
   ClientLeft      =   5865
   ClientTop       =   675
   ClientWidth     =   12675
   DrawStyle       =   2  'Dot
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":23FA
   ScaleHeight     =   10305
   ScaleWidth      =   12675
   ShowInTaskbar   =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtLogin 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   7065
      TabIndex        =   0
      Top             =   5940
      Width           =   1470
   End
   Begin VB.TextBox txtSenha 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   9435
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   5940
      Width           =   1035
   End
   Begin VB.Label lblResolucao 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "mensagem de resolucao"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5955
      TabIndex        =   4
      Top             =   9600
      Width           =   3495
   End
   Begin VB.Label lblSenha 
      AutoSize        =   -1  'True
      BackColor       =   &H0081E8FA&
      BackStyle       =   0  'Transparent
      Caption         =   "Senha"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   8670
      TabIndex        =   3
      Top             =   5985
      Width           =   675
   End
   Begin VB.Label lblLogin 
      AutoSize        =   -1  'True
      BackColor       =   &H0081E8FA&
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   6360
      TabIndex        =   2
      Top             =   5985
      Width           =   600
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub EfetuarLogin()

    Dim adologin As New ADODB.Recordset
    Dim sql As String
    
   On Error GoTo TrataErro
    
   sql = ("Select us_codigo,us_nome,US_NivelAcesso from glb_Usuariossistema where US_Nome ='" & txtLogin.Text & "' and US_senha = '" & _
      Me.txtSenha.Text & "'")
      adologin.CursorLocation = adUseClient
      adologin.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
      
   If adologin.EOF Then
      MsgBox "Usuário ou senha incorreto!", vbExclamation, "Falha no login"
      txtSenha = ""
   Else
      GLB_USU_Nome = UCase(Trim(adologin("Us_Nome")))
      GLB_USU_Codigo = Trim(adologin("US_Codigo"))
      GLB_USU_NivelAcesso = Trim(adologin("US_NivelAcesso"))

      frmBandeja.Show
      Unload frmLogin
   End If
   
   adologin.Close
   
TrataErro:
If Err.Number <> 0 Then
    Select Case Err.Number
        Case -2147467259
        MsgBox "Não foi possível verificar o usuário e senha" & vbNewLine & _
        "Verifique sua conexão com a rede", vbCritical, "Erro de conexão"
        sairDoSistema
        Case Else
        MsgBox "Erro ao verificar o usuário e senha" & vbNewLine & _
        "Verifique sua conexão com a rede", vbCritical, "Erro de conexão"
        sairDoSistema
    End Select
End If

End Sub



Private Sub Form_Load()

    'resolucaoOriginal.Colunas = resolucaoTela.Colunas
    'resolucaoOriginal.Linhas = resolucaoTela.Linhas
    'Call AlterarResolucao(1024, 768)

  left = (Screen.Width - Width) / 2
  top = (Screen.Height - Height) / 2
  
    If resolucaoTelaIdeal Then
        lblResolucao.Caption = ""
    Else
        lblResolucao.Caption = "A resolução " & resolucaoTela.Colunas & " x " & resolucaoTela.Linhas & " não é adequada" & vbNewLine _
        & "Resolução ideal: 1024 x 768"
    End If
  
End Sub

Private Sub txtSenha_GotFocus()
    campoSelecionadoComCaracter txtSenha
End Sub

Private Sub txtSenha_KeyPress(KeyAscii As Integer)
    KeyAscii = campoNormal(KeyAscii)
    If KeyAscii = 13 Then
        EfetuarLogin
    End If
    If KeyAscii = 27 Then
       sairDoSistema
    End If
End Sub

Private Sub txtlogin_GotFocus()
    campoSelecionadoComCaracter txtLogin
End Sub

Private Sub txtLogin_KeyPress(KeyAscii As Integer)
    proximoCampoEnter (KeyAscii)
    KeyAscii = campoNormal(KeyAscii)
    
    If KeyAscii = 27 Then
       sairDoSistema
    End If
End Sub
