Attribute VB_Name = "ModTrataErro"
Option Explicit

Public Sub erroBancoDeDados(numeroErro As ErrObject)
    If numeroErro.Number <> 0 Then
        Select Case numeroErro.Number
            Case -2147467259
            MsgBox "N�o foi poss�vel se conectar ao Banco de Dados" & vbNewLine & _
            "Verifique sua conex�o com a rede", vbCritical, Err.Source
        End Select
    End If
End Sub
