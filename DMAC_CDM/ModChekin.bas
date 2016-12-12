Attribute VB_Name = "ModChekin"
Option Explicit

Global ADO_Cn_Mobile As New ADODB.Connection
Global adoDemeoChekin As New ADODB.Recordset
Global adoDemeoChekinChekin2  As New ADODB.Recordset
Global adoDemeoChekin3 As New ADODB.Recordset
Global adoDemeoChekin4 As New ADODB.Recordset
Global adoDemeoChekin5 As New ADODB.Recordset
Global adoDemeoChekin6 As New ADODB.Recordset

Dim servidor, banco, Usuario, Senha As String

Sub inicioChekin()

    CarregarDBIni
    Call ConectaCDM
    
End Sub

Private Function CarregarDBIni()
    
  On Error GoTo erroConexaoBancoINI
    
  Dim ado_cn_dmac  As New ADODB.Connection
  Dim ADO_Cn_DmacA  As New ADODB.Connection
  Dim ADO_Cn_rsDmac As New ADODB.Recordset
    
  ADO_Cn_DmacA.Open "Driver={Microsoft Access Driver (*.mdb)};" & _
                    "Dbq=c:\sistemas\dmacini.mdb;" & _
                    "Uid=Admin; Pwd=astap36"
 
  sql = "Select * from ConexaoCDM"

  ADO_Cn_rsDmac.CursorLocation = adUseClient
  ADO_Cn_rsDmac.Open sql, ADO_Cn_DmacA, adOpenForwardOnly, adLockPessimistic
 
  With ADO_Cn_rsDmac
  
        If .BOF And .EOF Then
            MsgBox "Problemas no banco de dados de inicialização", vbCritical, "Erro"
            End
        Else
                        servidor = .Fields("TEF_Servidor")
                        banco = .Fields("TEF_Banco")
                        Usuario = .Fields("TEF_Usuario")
                        Senha = .Fields("TEF_Senha")
                        'Loja = .Fields("GLB_Loja")
        End If
        
    End With
     
    ADO_Cn_DmacA.Close
    
erroConexaoBancoINI:
    Select Case Err.Number
        Case -2147467259
        MsgBox "Não foi possível localizar ou conectar ao ini:" & vbNewLine _
        & "c:\sistemas\dmacini.mdb", vbCritical, "Erro"
        End
    End Select

End Function


Private Function ConectaCDM() As Boolean

    On Error GoTo ConexaoErro
    
    ADO_Cn_Mobile.Provider = "SQLOLEDB"
    ADO_Cn_Mobile.Properties("Data Source").Value = Nomeservidor
    ADO_Cn_Mobile.Properties("Initial Catalog").Value = BancoDeDados
    ADO_Cn_Mobile.Properties("User ID").Value = Usuario
    ADO_Cn_Mobile.Properties("Password").Value = Senha
    
    ADO_Cn_Mobile.Open
    
    ConectaCDM = True
    frmCheckin.Show
    
    Exit Function
        
ConexaoErro:
    
    ConectaCDM = False
        
    If Not Repetido("Login Failed") Then
        MsgBox "Erro na Conexão ADO", vbCritical, "Erro"
        
        MostraErro
    End If
     
End Function

