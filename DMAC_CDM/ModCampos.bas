Attribute VB_Name = "ModCampos"
Option Explicit

Public Function Replace(Texto As String, caracter As String, caracterParaSubstituir As String) As String
    
    Do While Texto Like "*" & caracter & "*"
        Texto = left$(Texto, (InStr(Texto, caracter) - 1)) _
        & caracterParaSubstituir _
        & right$(Texto, (Len(Texto) - (InStr(Texto, caracter))))
    Loop
    
    Replace = Texto
    
End Function

Public Sub campoSelecionadoComCaracter(campo As TextBox)
    If campo.Text <> "" Then
        campo.SelStart = 0
        campo.SelLength = Len(campo.Text)
    End If
End Sub

Public Function proximoCampoEnter(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ''SendKeys "{TAB}"
    End If
End Function

Public Function digitoLetra(KeyAscii As Integer) As Boolean
    digitoLetra = False
    If KeyAscii >= 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Then
        digitoLetra = True
    End If
End Function

Public Function digitoNumerico(KeyAscii As Integer) As Boolean
    digitoNumerico = False
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        digitoNumerico = True
    End If
End Function

Public Function digitoVirgulaPonto(KeyAscii As Integer) As Boolean
    digitoVirgulaPonto = False
    If KeyAscii = 44 Or KeyAscii = 46 Then
        digitoVirgulaPonto = True
    End If
End Function

Public Function digitoPadrao(KeyAscii As Integer) As Boolean
    digitoPadrao = False
    If KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 27 Then
        digitoPadrao = True
    End If
End Function

Public Function campoNumerico(KeyAscii As Integer) As Integer
    If digitoNumerico(KeyAscii) Or digitoPadrao(KeyAscii) Then
        campoNumerico = KeyAscii
    End If
End Function


Public Function campoNumericoVirgula(KeyAscii As Integer) As Integer
    If digitoNumerico(KeyAscii) Or digitoVirgulaPonto(KeyAscii) Or digitoPadrao(KeyAscii) Then
        campoNumericoVirgula = KeyAscii
    End If
End Function

Public Function campoNormal(KeyAscii As Integer) As Integer
    If digitoNumerico(KeyAscii) Or digitoLetra(KeyAscii) Or digitoPadrao(KeyAscii) Then
        campoNormal = KeyAscii
    End If
End Function

Public Function verificarCampoObrig(campo As TextBox) As Boolean
    
    If campo.Text = "" Then
        campo.SetFocus
        mensagemCampoObrigatorio campo.ToolTipText
        verificarCampoObrig = False
    Else
        verificarCampoObrig = True
    End If
    
End Function

Public Function validaCampoCNPJ(campo As TextBox) As Boolean
    If Len(campo.Text) = 14 Or Len(campo.Text) = 15 Then
            validaCampoCNPJ = True
        Else
            'mensagemCampoInvalido campo.ToolTipText
            campo.Text = ""
            validaCampoCNPJ = False
    End If
End Function

Public Sub autoPreencherZero(campo As TextBox)
    If campo.Text = "" Then
        campo.Text = "0,00"
    Else
        formataCampoDinheiro campo
    End If
End Sub

Public Sub formataCampoDinheiro(campoDinheiro As TextBox)
    campoDinheiro.Text = Format(campoDinheiro.Text, "##,#0.00")
End Sub

Public Function formataVariavelDinheiro(ByRef campoDinheiro As String)
    formataVariavelDinheiro = Format(campoDinheiro, "##,#0.00")
End Function

