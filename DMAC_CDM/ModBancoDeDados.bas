Attribute VB_Name = "ModBancoDeDados"
Option Explicit



Public Function buscaCodigoFornecedor(ByRef codigoOuCNPJ As String) As String
    Screen.MousePointer = 11

    Dim rdoFornecedor As New adodb.Recordset
    Dim sql As String
    
    If Len(codigoOuCNPJ) > 4 Then
        sql = "select FO_codigoFornecedor CodigoOuCNPJ from fornecedor where FO_CGC like '%" _
        & codigoOuCNPJ & "'"
    Else
        sql = "select FO_CGC CodigoOuCNPJ from fornecedor where FO_codigoFornecedor = '" _
        & codigoOuCNPJ & "'"
    End If
    
    rdoFornecedor.CursorLocation = adUseClient
    rdoFornecedor.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If rdoFornecedor.BOF And rdoFornecedor.EOF Then
        rdoFornecedor.Close
        'buscaNomeFornecedor = ""
        Exit Function
    Else
        buscaCodigoFornecedor = rdoFornecedor("CodigoOuCNPJ")
    End If

    rdoFornecedor.Close
    Screen.MousePointer = 0
End Function


Public Function buscaNomeFornecedor(ByRef codigoOuCNPJ As String) As String
    Screen.MousePointer = 11

    Dim rdoFornecedor As New adodb.Recordset
    Dim sql As String
    
    If Len(codigoOuCNPJ) > 4 Then
        sql = "select FO_NomeFantasia from fornecedor where FO_CGC like '%" _
        & codigoOuCNPJ & "'"
    Else
        sql = "select FO_NomeFantasia from fornecedor where FO_codigoFornecedor = '" _
        & codigoOuCNPJ & "'"
    End If
    
    rdoFornecedor.CursorLocation = adUseClient
    rdoFornecedor.Open sql, ADO_Cn_CDLocal, adOpenForwardOnly, adLockPessimistic
    
    If rdoFornecedor.BOF And rdoFornecedor.EOF Then
        rdoFornecedor.Close
        Exit Function
    Else
        buscaNomeFornecedor = rdoFornecedor("FO_NomeFantasia")
    End If

    rdoFornecedor.Close
    Screen.MousePointer = 0
End Function
