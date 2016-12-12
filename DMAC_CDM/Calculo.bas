Attribute VB_Name = "ModCalculo"
Option Explicit

Public Function calculaFrete(ByVal porcentagemFrete As Double, precoUnitario As Double, ByVal quantidade As Integer) As Double
    
    calculaFrete = ((porcentagemFrete * precoUnitario) / 100) * quantidade
    
End Function

Public Function calculaPorcentagem(ByVal Valor As Double, ValorTotal As Double) As Double

    If Valor <> 0 Then
        calculaPorcentagem = ((Valor * 100) / ValorTotal)
    Else
        calculaPorcentagem = 0
    End If
    
End Function

Public Function calculaDiferenca(maiorValor As Double, segundoValor As Double, desconto As Double) As Double

    If maiorValor = 0 Then
        calculaDiferenca = maiorValor
    Else
        calculaDiferenca = (maiorValor - segundoValor)
    End If
    
End Function

Public Function formatParaCalculo(ByVal Valor As String) As Double
    
    Valor = Replace(Valor, ".", "")
    If Valor = "" Then
        formatParaCalculo = 0
    Else
        formatParaCalculo = Valor
    End If
    
End Function

Public Function formatParaGravar(ByVal Valor As String) As String

    Valor = Replace(Valor, ".", "")
    Valor = Replace(Valor, ",", ".")
    formatParaGravar = Valor

End Function

