VERSION 5.00
Begin VB.Form frmListarArquivos 
   BackColor       =   &H00505050&
   BorderStyle     =   0  'None
   Caption         =   "Listar Arquivos"
   ClientHeight    =   4800
   ClientLeft      =   885
   ClientTop       =   930
   ClientWidth     =   11610
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4800
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   45
      Left            =   0
      ScaleHeight     =   45
      ScaleWidth      =   11565
      TabIndex        =   4
      Top             =   4080
      Width           =   11565
   End
   Begin VB.ComboBox cboArquivos 
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
      Left            =   135
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1920
      Width           =   11235
   End
   Begin VB.DirListBox dirPasta 
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
      Height          =   1440
      Left            =   2910
      TabIndex        =   1
      Top             =   90
      Width           =   8460
   End
   Begin VB.DriveListBox drvDrive 
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
      Left            =   150
      TabIndex        =   0
      Top             =   105
      Width           =   2520
   End
   Begin CentroDeDistribuicao.chameleonButton cmdListar 
      Height          =   510
      Left            =   8730
      TabIndex        =   5
      Top             =   4215
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "Listar"
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
      MICON           =   "frmListarArquivos.frx":0000
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
      Left            =   10170
      TabIndex        =   6
      Top             =   4215
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "OK"
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
      MICON           =   "frmListarArquivos.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblCaminho 
      BackColor       =   &H00505050&
      Caption         =   "label1"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   180
      TabIndex        =   3
      Top             =   2610
      Visible         =   0   'False
      Width           =   5805
   End
End
Attribute VB_Name = "frmListarArquivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private m_strPasta As String
Dim liberaBotaoOk As Boolean

Public Sub ListarArquivosDiretorio()
    On Error GoTo Handle_Error

          Dim objFSO                  As Scripting.FileSystemObject
          Dim objPasta                As Object
          Dim objArquivo              As Object
          Dim objArquivosExistentes   As Object
          Dim strNomeArquivo          As String
    
1        Set objFSO = New Scripting.FileSystemObject
2        Set objPasta = objFSO.GetFolder(m_strPasta)
3        Set objArquivosExistentes = objPasta.Files
    
4        cboArquivos.Clear
    
5        For Each objArquivo In objArquivosExistentes
6            Call cboArquivos.AddItem(objArquivo.Name)
7        Next
        
         If cboArquivos.ListCount > 0 Then
            cboArquivos.ListIndex = 0
         End If
    
8        Exit Sub

Handle_Error:
    Debug.Print "Número: " & Err.Number & vbCrLf & "Descrição: " & Err.Description & vbCrLf & "Linha: " & Erl & vbCrLf
End Sub



Private Sub cboArquivos_Click()
    lblCaminho.Caption = m_strPasta & "\" & cboArquivos.Text
    If InStr(lblCaminho.Caption, "c:\teste xml\") = 0 Then
       MsgBox "Caminho inválido!" & vbCrLf & "Selecione um arquivo XML do diretório: c:\teste xml\", vbCritical, "ATENÇÃO"
       liberaBotaoOk = False
    Else
       liberaBotaoOk = True
    End If
End Sub

Private Sub cmdListar_Click()
    Call ListarArquivosDiretorio
End Sub

Private Sub cmdOK_Click()
 If liberaBotaoOk = True Then
    frmConverterXML.Text1.Text = Me.lblCaminho
    
    frmConverterXML.Show 1
    Unload Me
  ' frmXml.Show 1
 End If
End Sub

Private Sub cmdRetorna_Click()
    Unload Me
End Sub

Private Sub dirPasta_Change()
    m_strPasta = dirPasta.Path
End Sub

Private Sub drvDrive_Change()
    On Error Resume Next
    dirPasta.Path = left(drvDrive.Drive, 2) & "\"
    
    If Err.Number = 76 Then 'Erro: Path not found
        MsgBox "Diretório (drive) não está disponível", vbCritical, "Aviso"
    End If
End Sub

Private Sub Form_Load()
    m_strPasta = App.Path
    dirPasta.Path = drvDrive.Drive & "\"
        
   ' frmListarArquivos.top = (Screen.Height - frmListarArquivos.Height) / 2
   ' frmListarArquivos.left = (Screen.Width - frmListarArquivos.Width) / 2
   
    frmListarArquivos.top = 5700
    frmListarArquivos.left = 90
    
   liberaBotaoOk = False
   
   telaChamou = ""
End Sub

