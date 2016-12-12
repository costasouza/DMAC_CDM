Attribute VB_Name = "ModTrataErro"
Option Explicit

Public Sub erroBancoDeDados(numeroErro As ErrObject)
    If numeroErro.Number <> 0 Then
        Select Case numeroErro.Number
            Case -2147467259
            MsgBox "Não foi possível se conectar ao Banco de Dados" & vbNewLine & _
            "Verifique sua conexão com a rede", vbCritical, Err.Source
        End Select
    End If
End Sub
