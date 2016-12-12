Attribute VB_Name = "ModMensagens"
Option Explicit

Public Sub mensagemCampoObrigatorio(NomeCampo As String)
    MsgBox NomeCampo & " � um campo obrigatorio!", vbInformation, "Campo obrigatorio n�o preenchido"
End Sub


Public Sub mensagemCampoInvalido(NomeCampo As String)
    MsgBox NomeCampo & " inv�lido!", vbExclamation, "Campo inv�lido"
End Sub


Public Function mensagemLimparCampos() As Boolean

    If MsgBox("Deseja limpar todos os campos?", vbQuestion + vbYesNo, "Limpar todos os campos") = vbYes Then
            mensagemLimparCampos = True
    Else
            mensagemLimparCampos = False
    End If
    
End Function


Public Function mensagemExluir(NomeCampo As String) As Boolean

    If MsgBox("Deseja exluir " & NomeCampo & "?", vbQuestion + vbYesNo, "Excluir") = vbYes Then
            mensagemExluir = True
    Else
            mensagemExluir = False
    End If
    
End Function



