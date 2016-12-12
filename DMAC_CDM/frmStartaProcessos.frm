VERSION 5.00
Begin VB.Form frmStartaProcessos 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Starta Processos"
   ClientHeight    =   5925
   ClientLeft      =   120
   ClientTop       =   3810
   ClientWidth     =   15120
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5925
   ScaleMode       =   0  'User
   ScaleWidth      =   15120
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPedido 
      Height          =   285
      Left            =   255
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   300
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Timer Timer1 
      Left            =   420
      Top             =   3780
   End
End
Attribute VB_Name = "frmStartaProcessos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
 
    Call carregarPosicaoTamanhoTela(Me)
    
    Me.top = frmControleCD.lblNomeTelas.top
    Me.Height = Me.Height + frmControleCD.lblNomeTelas.Height + 15
    
    Screen.MousePointer = 11
    
    frmStartaProcessos.Picture = LoadPicture(endIMG("FundoProcessa"))
    If GLB_modoOffline = False Then Call StatusAtualizacao
    Esperar 3
    
    Screen.MousePointer = 0
    Unload Me

End Sub

Private Sub StatusAtualizacao()

    Screen.MousePointer = 11
    
    Dim adoSerieNotaFiscal As New ADODB.Recordset
    
    'sql = "Select vc_serie as serie from capanfvenda where vc_serie = 'NE' and vc_notafiscal = " & NroNotaFiscal
         'adoSerieNotaFiscal.CursorLocation = adUseClient
         'adoSerieNotaFiscal.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    'If Not adoSerieNotaFiscal.EOF Then
    
                   'sql = "exec SP_VDA_Cria_NFe_CD '" & lojaorigem & "'," & NroNotaFiscal & ",'NE','" & Carimbo & "'"
                   'ADO_Cn_CDLocal.Execute (sql)
                   
    'End If
    
    'adoSerieNotaFiscal.Close
    
    sql = "exec SP_Atualiza_Processos_Venda_Central"
    ADO_Cn_CDLocal.Execute sql
    
    Screen.MousePointer = 0

End Sub


Sub Esperar(ByVal Tempo As Integer)
    Dim StartTime As Long
    StartTime = Timer
    Do While Timer < StartTime + Tempo
       DoEvents
    Loop
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Imagemfundo_Click()

End Sub

