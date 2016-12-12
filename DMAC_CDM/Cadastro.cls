VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Public Conexao As Variant
Public UseWhere As Boolean
Public ControlFields As Boolean
Public Adicionar As Boolean
Public IncludeExtraFields As Boolean
Public ClausulaWhere As String
Public ExecProcedure As Boolean

Dim WmanuProduto As Boolean
Dim HoraManutencao As String

Dim ConectouAfgOK As Boolean

Dim matControles() As String
Dim matCampos() As String
Dim matExtraFields() As String
Dim matExtraValues() As String
Dim matTipoFormato() As Integer
Dim matValorLimpeza() As String
Dim matDados() As String
Dim matFazerVerificacao() As String
Dim Formulario As Form
Dim ID_Form As Long
Dim DisparaEventos As Boolean

Dim Registro As rdoResultset
Dim RsFornecedorAfg As rdoResultset
Dim WgravaAfg As Boolean

Private Const Moeda = 1
Private Const numero = 2
Private Const Caractere = 3
Private Const data = 4
Private Const vItemData = 5
Private Const Valor = 6
Private Const TotalText = 7
Private Const vIndice = 8
Private Const RightText = 9
Private Const Porcentagem = 10
Private Const LeftText = 11
Private Const vDecimal = 12
Private Const CPF = 13
Private Const CEP = 14

Private mvarSQL As String
Private mvarControles As String
Private mvarCampos As String
Private mvarExtraFields As String
Private mvarExtraValues As String
Private mvarTipoFormato As String
Private mvarValorLimpeza As String
Private mvarFazerVerificacao As String
Private mvarNomeFormulario As String

Public Event RegistroNaoExiste()
Public Event RegistroEncontrado(ByRef rdoInterno As Variant, ByRef Cancelar As Boolean)
Public Event LeituraOK()
Public Event LeituraErro(ByVal ErroNumero As Long, ByVal Descricao As String)
Public Event PreenchimentoOK()
Public Event PreenchimentoErro(ByVal ErroNumero As Long, ByVal Descricao As String, ByVal NomeControle As String)
Public Event GravacaoOK(ByVal resultado As String)
Public Event GravacaoErro(ByVal ErroNumero As Long, ByVal Descricao As String)
Public Event VerificacaoErro(ByVal Controle As String)
Public Event VerificacaoOK()

Public Property Let NomeFormulario(ByVal vData As String)
    
    Dim Indice As Long
    Dim Maximo As Long
    
    Maximo = Forms.Count - 1
    
    mvarNomeFormulario = ""
    
    For Indice = 0 To Maximo Step 1
        If UCase(Forms(Indice).Name) = UCase(vData) Then
            mvarNomeFormulario = vData
            Set Formulario = Forms(Indice)
            Exit For
        End If
    Next Indice
    
    If mvarNomeFormulario = "" Then
        Err.Raise -10, "Classe Cadastro", "Nome de formulario não encontrado."
    End If

End Property

Public Property Get NomeFormulario() As String
    
    NomeFormulario = mvarNomeFormulario

End Property

Public Property Let FazerVerificacao(ByVal vData As String)
    
    mvarFazerVerificacao = vData
    
    DisparaEventos = True
    
    MontarMatriz vData, matFazerVerificacao

End Property

Public Property Get FazerVerificacao() As String
    
    FazerVerificacao = mvarFazerVerificacao

End Property

Public Property Let ValorLimpeza(ByVal vData As String)
    
    mvarValorLimpeza = vData
    MontarMatriz vData, matValorLimpeza

End Property

Public Property Get ValorLimpeza() As String
    
    ValorLimpeza = mvarValorLimpeza

End Property

Public Property Let TipoFormato(ByVal vData As String)
    
    mvarTipoFormato = vData
    
    Dim Tamanho As Long
    Dim CharPos As Long
    Dim NomeLido As String
    Dim Indice As Long
    Dim PosAnt As Long
    
    Tamanho = Len(vData)
    
    ReDim matTipoFormato(0) As Integer
    
    If Tamanho > 0 Then
        CharPos = InStr(vData, ";")
        If CharPos = 0 Then
            matTipoFormato(0) = vData
        Else
            PosAnt = 1
            Indice = 0
            Do While CharPos <> 0
                NomeLido = UCase(Mid(vData, PosAnt, CharPos - PosAnt))
                PosAnt = CharPos + 2
                CharPos = InStr(CharPos + 1, vData, ";")
                
                VerFormato NomeLido, Indice
                
                Indice = Indice + 1
                ReDim Preserve matTipoFormato(Indice) As Integer
            Loop
            NomeLido = UCase(Mid(vData, PosAnt))
            VerFormato NomeLido, Indice
        End If
    End If

End Property

Private Sub VerFormato(ByVal NomeLido As String, ByVal Indice As Long)

    Select Case NomeLido
        Case "MOEDA"
            matTipoFormato(Indice) = Moeda
        Case "NUMERO"
            matTipoFormato(Indice) = numero
        Case "CARACTERE"
            matTipoFormato(Indice) = Caractere
        Case "DATA"
            matTipoFormato(Indice) = data
        Case "ITEMDATA"
            matTipoFormato(Indice) = vItemData
        Case "VALOR"
            matTipoFormato(Indice) = Valor
        Case "TOTALTEXT"
            matTipoFormato(Indice) = TotalText
        Case "INDICE"
            matTipoFormato(Indice) = vIndice
        Case "RIGHTTEXT"
            matTipoFormato(Indice) = RightText
        Case "PORCENTAGEM"
            matTipoFormato(Indice) = Porcentagem
        Case "LEFTTEXT"
            matTipoFormato(Indice) = LeftText
        Case "DECIMAL"
            matTipoFormato(Indice) = vDecimal
        Case "CPF"
            matTipoFormato(Indice) = CPF
        Case "CEP"
            matTipoFormato(Indice) = CEP
    End Select

End Sub

Public Property Get TipoFormato() As String
    
    TipoFormato = mvarTipoFormato

End Property

Public Property Let ExtraValues(ByVal vData As String)
    
    mvarExtraValues = vData
    MontarMatriz vData, matExtraValues
    
End Property

Public Property Get ExtraValues() As String
    
    ExtraValues = mvarExtraValues

End Property

Public Property Let ExtraFields(ByVal vData As String)
    
    mvarExtraFields = vData
    MontarMatriz vData, matExtraFields

End Property

Public Property Get ExtraFields() As String
    
    ExtraFields = mvarExtraFields

End Property

Public Sub Re_Get()

    DisparaEventos = False
    
    Verificar
    
    DisparaEventos = True

End Sub

Public Sub Verificar()

    Dim Indice As Long
    Dim Maximo As Long
    
    If DisparaEventos Then
        On Error GoTo ErroVerificacao
    Else
        On Error Resume Next
    End If
    
    Maximo = UBound(matControles)
    
    ReDim Preserve matDados(Maximo) As String
    
    For Indice = 0 To Maximo Step 1
        Select Case matTipoFormato(Indice)
            Case Moeda
                matDados(Indice) = Decimals(Indice)
            Case numero
                matDados(Indice) = Numeros2(Formulario.Controls(matControles(Indice)))
            Case Caractere
                matDados(Indice) = Trim(Formulario.Controls(matControles(Indice)))
            Case data
                matDados(Indice) = VerificaData(Indice)
            Case vItemData
                matDados(Indice) = IIf(Formulario.Controls(matControles(Indice)) <> "", Formulario.Controls(matControles(Indice)).ItemData(Formulario.Controls(matControles(Indice)).ListIndex), "")
            Case Valor
                If TypeOf Formulario.Controls(matControles(Indice)) Is ComboBox Then
                    matDados(Indice) = Procura(Indice, Valor)
                Else
                    matDados(Indice) = Val(Formulario.Controls(matControles(Indice)))
                End If
            Case TotalText
                matDados(Indice) = Procura(Indice, TotalText)
            Case vIndice
                matDados(Indice) = IIf(Formulario.Controls(matControles(Indice)).ListIndex >= 0, Formulario.Controls(matControles(Indice)).ListIndex, "")
            Case RightText
                matDados(Indice) = Procura(Indice, RightText)
            Case Porcentagem, vDecimal
                matDados(Indice) = Decimals(Indice)
            Case LeftText
                matDados(Indice) = Procura(Indice, LeftText)
            Case CPF
                matDados(Indice) = VerificaCPF(Formulario.Controls(matControles(Indice)))
            Case CEP
                matDados(Indice) = VerificaCEP(Formulario.Controls(matControles(Indice)))
        End Select
    Next Indice
    
    If DisparaEventos Then
        For Indice = 0 To Maximo Step 1
            If Val(matFazerVerificacao(Indice)) = 1 And matDados(Indice) = "" Then
                RaiseEvent VerificacaoErro(matControles(Indice))
                Exit Sub
            End If
        Next Indice
        
        RaiseEvent VerificacaoOK
    End If
    
    DisparaEventos = True
    
    Exit Sub
    
ErroVerificacao:
    RaiseEvent VerificacaoErro("Erro " & Err.Number & ": " & Err.Description)

End Sub

Private Function VerificaCPF(ByVal Texto As String) As String

    Texto = Numeros2(Texto)
    
    Select Case Len(Texto)
        Case Is < 11
            VerificaCPF = ""
            Exit Function
        Case Is > 11
            Texto = left(Texto, 11)
    End Select
    
    VerificaCPF = Texto

End Function

Private Function VerificaCEP(ByVal Texto As String) As String

    Texto = Numeros2(Texto)
    
    Select Case Len(Texto)
        Case Is < 8
            VerificaCEP = ""
            Exit Function
        Case Is > 8
            Texto = left(Texto, 8)
    End Select
    
    VerificaCEP = Texto

End Function

Private Function FormataCPF(ByVal Texto As String) As String

    Dim Retorno As String
    
    Texto = Numeros2(Texto)
    
    Select Case Len(Texto)
        Case Is < 11
            FormataCPF = ""
            Exit Function
        Case Is > 11
            Texto = left(Texto, 11)
    End Select
    
    Retorno = left(Texto, 3) & "." & Mid(Texto, 4, 3) & "." & Mid(Texto, 7, 3) & "-" & right(Texto, 2)
    
    FormataCPF = Retorno

End Function

Private Function FormataCEP(ByVal Texto As String) As String

    Dim Retorno As String

    Texto = Numeros2(Texto)
    
    Select Case Len(Texto)
        Case Is < 8
            FormataCEP = ""
            Exit Function
        Case Is > 8
            Texto = left(Texto, 8)
    End Select
    
    Retorno = left(Texto, 5) & "-" & right(Texto, 3)
    
    FormataCEP = Retorno

End Function

Private Function Procura(ByVal Indice As Long, ByVal O_Que As Long) As String

    Dim varApoio As String
    Dim Ponteiro As Long
    Dim Maximo As Long
    Dim Retorno As Boolean
    
    If Val(matFazerVerificacao(Indice)) = 0 Then
        Select Case O_Que
            Case Valor
                Procura = 0
            Case TotalText
                Procura = ""
            Case RightText
                Procura = ""
            Case LeftText
                Procura = ""
        End Select
        Exit Function
    End If
    
    Maximo = Formulario.Controls(matControles(Indice)).ListCount - 1
    
    Procura = ""
    Retorno = False
    If Maximo >= 0 Then
        Select Case O_Que
            Case Valor
                varApoio = Val(Formulario.Controls(matControles(Indice)))
            Case TotalText
                varApoio = UCase(Formulario.Controls(matControles(Indice)))
            Case RightText
                If InStr(Formulario.Controls(matControles(Indice)), "-") > 0 Then
                    varApoio = Trim(right(Formulario.Controls(matControles(Indice)), Len(Formulario.Controls(matControles(Indice))) - (InStr(Formulario.Controls(matControles(Indice)), "-") + 1)))
                Else
                    varApoio = Formulario.Controls(matControles(Indice))
                End If
            Case LeftText
                If InStr(Formulario.Controls(matControles(Indice)), "-") > 0 Then
                    varApoio = Trim(left(Formulario.Controls(matControles(Indice)), (InStr(Formulario.Controls(matControles(Indice)), "-") - 1)))
                Else
                    varApoio = Formulario.Controls(matControles(Indice))
                End If
        End Select
        
        For Ponteiro = 0 To Maximo Step 1
            Select Case O_Que
                Case Valor
                    Retorno = IIf(varApoio = Val(Formulario.Controls(matControles(Indice)).List(Ponteiro)), True, False)
                Case TotalText
                    Retorno = IIf(varApoio = UCase(Formulario.Controls(matControles(Indice)).List(Ponteiro)), True, False)
                Case RightText
                    Retorno = IIf(varApoio = Trim(right(Formulario.Controls(matControles(Indice)).List(Ponteiro), Len(Formulario.Controls(matControles(Indice)).List(Ponteiro)) - (InStr(Formulario.Controls(matControles(Indice)).List(Ponteiro), "-") + 1))), True, False)
                Case LeftText
                    Retorno = IIf(varApoio = Trim(left(Formulario.Controls(matControles(Indice)).List(Ponteiro), (InStr(Formulario.Controls(matControles(Indice)).List(Ponteiro), "-") - 1))), True, False)
            End Select
            If Retorno Then
                Procura = varApoio
                Formulario.Controls(matControles(Indice)).ListIndex = Ponteiro
                Exit Function
            End If
        Next Ponteiro
    End If
    Formulario.Controls(matControles(Indice)).ListIndex = -1

End Function

Private Function VerificaData(ByVal Indice As Long) As String

    On Error Resume Next
    
    Err.Clear
    
    If IsDate(Formulario.Controls(matControles(Indice))) Then
        If Err.Number = 0 Then
            VerificaData = Formulario.Controls(matControles(Indice))
        Else
            Err.Clear
            VerificaData = ""
        End If
    Else
        VerificaData = ""
    End If

End Function

Private Function Decimals(ByVal Indice As Long) As String

    Dim varApoio As String
    Dim Tamanho As Long
    
    varApoio = Trim(Formulario.Controls(matControles(Indice)))
    
    If varApoio = "" Then
        Decimals = ""
        Exit Function
    End If
    
    varApoio = Format(varApoio, "0.00")
    Tamanho = Len(varApoio)
    If IsNumeric(varApoio) Then
        varApoio = left(varApoio, Tamanho - 3) & "." & right(varApoio, 2)
    Else
        varApoio = ""
    End If
    
    Decimals = varApoio

End Function

Public Sub Limpar()

    Dim Indice As Long
    Dim Maximo As Long
    
    On Error GoTo Sair
    
    Maximo = UBound(matControles)
    
    For Indice = 0 To Maximo Step 1
        If LCase(left(matControles(Indice), 3)) = "cmb" Then
            If Formulario.Controls(matControles(Indice)).Style = 2 Then
                Formulario.Controls(matControles(Indice)).ListIndex = Trim(matValorLimpeza(Indice))
            Else
                Formulario.Controls(matControles(Indice)) = Trim(matValorLimpeza(Indice))
            End If
        Else
            Formulario.Controls(matControles(Indice)) = Trim(matValorLimpeza(Indice))
        End If
    Next Indice
    
Sair:
    Err.Clear

End Sub

Public Sub Preencher()

    Dim Indice As Long
    Dim Maximo As Long
    Dim Retorno As Long
    
    On Error GoTo PreencheErro
    
    Indice = 0
    
    Maximo = UBound(matTipoFormato)
    
    For Indice = 0 To Maximo Step 1
        Select Case matTipoFormato(Indice)
            Case Moeda
                Formulario.Controls(matControles(Indice)) = Format(CDbl(matDados(Indice)), "R$ ###,###,###,###,###0.00")
                
            Case numero, Caractere
                Formulario.Controls(matControles(Indice)) = matDados(Indice)
                
            Case Porcentagem
                Formulario.Controls(matControles(Indice)) = Format(CDbl(matDados(Indice)), "0.00") & "%"
                
            Case data
                Formulario.Controls(matControles(Indice)) = Format(matDados(Indice), "dd/mm/yyyy")
                
            Case vItemData
                Formulario.Controls(matControles(Indice)).ListIndex = LocalizaItemData(Indice)
                
            Case Valor
                If TypeOf Formulario.Controls(matControles(Indice)) Is ComboBox Then
                    Retorno = LocalizaValor(Indice)
                    Formulario.Controls(matControles(Indice)).ListIndex = Retorno
                    If Retorno = -1 And Formulario.Controls(matControles(Indice)).Style = 0 Then
                        Formulario.Controls(matControles(Indice)) = matDados(Indice)
                    End If
                    
                Else
                    Formulario.Controls(matControles(Indice)) = matDados(Indice)
                End If
                
            Case TotalText
                Retorno = LocalizaTotalText(Indice)
                Formulario.Controls(matControles(Indice)).ListIndex = Retorno
                If Retorno = -1 And Formulario.Controls(matControles(Indice)).Style = 0 Then
                    Formulario.Controls(matControles(Indice)) = matDados(Indice)
                End If

            Case vIndice
                Formulario.Controls(matControles(Indice)).ListIndex = Val(matDados(Indice))
                
            Case RightText
                Retorno = LocalizaRightText(Indice)
                Formulario.Controls(matControles(Indice)).ListIndex = Retorno
                If Retorno = -1 And Formulario.Controls(matControles(Indice)).Style = 0 Then
                    Formulario.Controls(matControles(Indice)) = matDados(Indice)
                End If
            
            Case LeftText
                Retorno = LocalizaLeftText(Indice)
                Formulario.Controls(matControles(Indice)).ListIndex = Retorno
                If Retorno = -1 And Formulario.Controls(matControles(Indice)).Style = 0 Then
                    Formulario.Controls(matControles(Indice)) = matDados(Indice)
                End If
                
            Case vDecimal
                Formulario.Controls(matControles(Indice)) = Format(CDbl(matDados(Indice)), "0.00")
                
            Case CPF
                Formulario.Controls(matControles(Indice)) = FormataCPF(matDados(Indice))
            
            Case CEP
                Formulario.Controls(matControles(Indice)) = FormataCEP(matDados(Indice))
                
        End Select
    Next Indice
    
    RaiseEvent PreenchimentoOK
    
    Exit Sub
    
PreencheErro:

    RaiseEvent PreenchimentoErro(Err.Number, Err.Description, matControles(Indice))
    Err.Clear

End Sub

Public Sub Gravar()

    Dim SQLFinal As String
    Dim SQLFinalAfg As String
    Dim CamposLidos As String
    Dim CamposLidosAfg As String
    Dim ValoresLidos As String
    Dim IniciouTransacao As Boolean

    IniciouTransacao = False
    
    Err.Clear
    On Error Resume Next
    
    SQLFinal = ""
    CamposLidos = ""
    ValoresLidos = ""
    
    If ExecProcedure Then
        ExecutarProcedure mvarSQL & " " & Parametros
        Exit Sub
    End If
    
    If Adicionar Then
        CamposLidos = LerCampos(matCampos)
        ValoresLidos = LerValores(1)
        If IncludeExtraFields Then
            CamposLidos = CamposLidos & ", " & LerCampos(matExtraFields)
            ValoresLidos = ValoresLidos & ", " & LerValores(2)
        End If
        
        SQLFinal = "Insert Into " & mvarSQL & " (" & CamposLidos & ") Values (" & ValoresLidos & ")"
        SQLFinalAfg = "Insert Into " & mvarSQL & " (" & CamposLidos & ") Values (" & ValoresLidos & ")"
    Else
        If ClausulaWhere <> "" Then
            CamposLidos = ObterUpdateValores
            CamposLidosAfg = ObterUpdateValoresAfg
            If WmanuProduto = True Then
               CamposLidos = CamposLidos & ",Pr_HoraManutencao='" & HoraManutencao & "'"
               CamposLidosAfg = CamposLidosAfg & ",Pr_HoraManutencao='" & HoraManutencao & "'"
               WmanuProduto = False
            End If
            SQLFinal = "Update " & mvarSQL & " Set " & CamposLidos & " " & ClausulaWhere
            SQLFinalAfg = "Update " & mvarSQL & " Set " & CamposLidosAfg & " " & ClausulaWhere
        Else
            Exit Sub
        End If
    End If
    
    
    Conexao.BeginTrans
    IniciouTransacao = True
    Conexao.Execute SQLFinal
    
    If Err.Number = 0 Then
       Conexao.CommitTrans
    Else
       If IniciouTransacao Then
          Conexao.RollbackTrans
       End If
       RaiseEvent GravacaoErro(Err.Number, Err.Description)
       Err.Clear
       Exit Sub
    End If
    
    ConectouAfgOK = False
       
'    If NomeServidor <> "Retaguarda" Then
'        If WmanuProduto = True Then
'           ConectaODBCBatchAFGP "", "", ""
'        Else
'           ConectouAfgOK = False
'           WmanuProduto = False
'        End If
'
'        If ConectouAfgOK = True Then
'           If Wfornecedor > 0 Then
'              SQL = "Select FO_CodigoFornecedor from Fornecedor " _
'                  & "where FO_CodigoFornecedor= " & Wfornecedor & " " _
'                  & "and FO_TipoFornecedor = 'F'"
'
'              Set RsFornecedorAfg = rdoCnSupBatch.OpenResultset(SQL)
'
'              If RsFornecedorAfg.EOF Then
'                 MsgBox "Fornecedor não cadastrado Na AFG, Esta operação não sera efetuada na Afg", vbInformation, "Atenção"
'                 ConectouAfgOK = False
'                 RsFornecedorAfg.Close
'                 ConexaoAFG.Close
'              End If
'           End If
'
'        End If
'
'        If ConectouAfgOK = True Then
'           Status "Aguarde gravando produto AFG"
'           ConexaoAFG.BeginTrans
'           ConexaoAFG.Execute SQLFinalAfg
'
'           If Err.Number = 0 Then
'               ConexaoAFG.CommitTrans
'               Status "Pronto                                      "
'           Else
'               ConexaoAFG.RollbackTrans
'               Status "Pronto                                      "
'               MsgBox "Não foi possivel gravar produto na AFG, Favor tentar novamente", vbInformation, "Atenção"
'               ConexaoAFG.Close
'               ConectouAfgOK = False
'           End If
'        End If
'
'    End If
    
    RaiseEvent GravacaoOK("0")

'ErroGravacao:
'
'    If IniciouTransacao Then
'        Conexao.RollbackTrans
'    End If
'
'    RaiseEvent GravacaoErro(Err.Number, Err.Description)
'
'    Err.Clear

End Sub

Private Function Parametros() As String

    Dim Retorno As String

    Retorno = ""
    
    Retorno = LerValores(1)
    
    If IncludeExtraFields Then
        Retorno = Retorno & ", " & LerValores(2)
    End If
    
    Parametros = Retorno
    
End Function

Private Sub ExecutarProcedure(ByVal SQL1 As String)

    Dim rdoProc As rdoResultset
    Dim resultado As String
    Dim Maximo As Long
    Dim Indice As Long

    On Error Resume Next
    
    Set rdoProc = Conexao.OpenResultset(SQL1, Options:=rdExecDirect)
    
    If Err.Number = 0 Then
        Maximo = rdoProc.rdoColumns.Count - 1
        resultado = ""
        For Indice = 0 To Maximo Step 1
            resultado = resultado & "; " & CStr(rdoProc(Indice))
        Next Indice
        
        rdoProc.Close
        
        RaiseEvent GravacaoOK(Mid(resultado, 3))
    Else
        rdoProc.Close
        RaiseEvent GravacaoErro("-1", "Não foi possível gravar/atualizar registro.")
    End If

End Sub

Private Function ObterUpdateValores() As String

    Dim Indice As Long
    Dim Maximo As Long
    Dim Retorno As String
    
    Maximo = UBound(matCampos)
    Retorno = ""
    For Indice = 0 To Maximo Step 1
           Retorno = Retorno & matCampos(Indice) & " = " & TipoValorSQL(matDados(Indice), matTipoFormato(Indice)) & ", "
    Next Indice
    
    If Retorno <> "" Then
        Retorno = left(Retorno, Len(Retorno) - 2)
    End If
    
    ObterUpdateValores = Retorno

End Function


Private Function ObterUpdateValoresAfg() As String

    Dim Indice As Long
    Dim Maximo As Long
    Dim Retorno As String
    
    Maximo = UBound(matCampos)
    Retorno = ""
    For Indice = 0 To Maximo Step 1
        If matCampos(Indice) <> "PR_Markup" And matCampos(Indice) <> "PR_PrecoCusto1" And matCampos(Indice) <> "PR_PrecoVenda1" And matCampos(Indice) <> "PR_DATACADASTRO" And matCampos(Indice) <> "PR_PRECOFORNECEDOR" Then
           Retorno = Retorno & matCampos(Indice) & " = " & TipoValorSQL(matDados(Indice), matTipoFormato(Indice)) & ", "
        End If
    Next Indice
    
    If Retorno <> "" Then
        Retorno = left(Retorno, Len(Retorno) - 2)
    End If
    
    ObterUpdateValoresAfg = Retorno

End Function



Private Function LerValores(ByVal Tipo As Long) As String

    Dim Indice As Long
    Dim Maximo As Long
    Dim Retorno As String
    
    If Tipo = 1 Then
        Maximo = UBound(matDados)
        Retorno = ""
        For Indice = 0 To Maximo Step 1
            Retorno = Retorno & TipoValorSQL(matDados(Indice), matTipoFormato(Indice)) & ", "
        Next Indice
    Else
        Maximo = UBound(matExtraValues)
        Retorno = ""
        For Indice = 0 To Maximo Step 1
            Retorno = Retorno & matExtraValues(Indice) & ", "
        Next Indice
    End If
    
    If Retorno <> "" Then
        Retorno = left(Retorno, Len(Retorno) - 2)
    End If

    LerValores = Retorno

End Function

Private Function TipoValorSQL(ByVal Valor As String, ByVal Tipo As Long) As String

    Select Case Tipo
'        Case Moeda, Numero, Porcentagem, vItemData, Valor, vIndice, vDecimal
'            Valor = Valor
        Case Caractere, TotalText, RightText, LeftText, CPF, CEP
            Valor = "'" & Valor & "'"
        Case data
            Valor = "'" & Mid(Valor, 4, 3) & left(Valor, 2) & right(Format(Valor, "dd/mm/yyyy"), 5) & "'"
    End Select
    
    TipoValorSQL = Valor

End Function

'Public Sub Validar()
'
'    Dim Indice As Long
'    Dim Maximo As Long
'    Dim Retorno As Long
'
'    On Error GoTo ValidacaoErro
'
'    Indice = 0
'
'    Maximo = UBound(matTipoFormato)
'
'    For Indice = 0 To Maximo Step 1
'        Select Case matTipoFormato(Indice)
'            Case Moeda
'                Formulario.Controls(matControles(Indice)) = Format(CDbl(matDados(Indice)), "R$ ###,###,###,###,###0.00")
'
'            Case Numero, Caractere
'                Formulario.Controls(matControles(Indice)) = matDados(Indice)
'
'            Case Porcentagem
'                Formulario.Controls(matControles(Indice)) = Format(CDbl(matDados(Indice)), "0.00") & "%"
'
'            Case Data
'                Formulario.Controls(matControles(Indice)) = Format(matDados(Indice), "dd/mm/yyyy")
'
'            Case vItemData
'                Formulario.Controls(matControles(Indice)).ListIndex = LocalizaItemData(Indice)
'
'            Case Valor
'                Retorno = LocalizaValor(Indice)
'                Formulario.Controls(matControles(Indice)).ListIndex = Retorno
'                If Retorno = -1 And Formulario.Controls(matControles(Indice)).Style = 0 Then
'                    Formulario.Controls(matControles(Indice)) = matDados(Indice)
'                End If
'
'            Case TotalText
'                Retorno = LocalizaTotalText(Indice)
'                Formulario.Controls(matControles(Indice)).ListIndex = Retorno
'                If Retorno = -1 And Formulario.Controls(matControles(Indice)).Style = 0 Then
'                    Formulario.Controls(matControles(Indice)) = matDados(Indice)
'                End If
'
'            Case vIndice
'                Formulario.Controls(matControles(Indice)).ListIndex = Val(matDados(Indice))
'
'            Case RightText
'                Retorno = LocalizaRightText(Indice)
'                Formulario.Controls(matControles(Indice)).ListIndex = Retorno
'                If Retorno = -1 And Formulario.Controls(matControles(Indice)).Style = 0 Then
'                    Formulario.Controls(matControles(Indice)) = matDados(Indice)
'                End If
'
'        End Select
'    Next Indice
'
'    RaiseEvent PreenchimentoOK
'
'    Exit Sub
'
'ValidacaoErro:
'
'    RaiseEvent PreenchimentoErro(Err.Number, Err.Description, matControles(Indice))
'    Err.Clear
'
'End Sub

Public Sub Ler()

    Dim NovaSQL As String
    Dim Indice As Long
    Dim Maximo As Long
    Dim Cancelou As Boolean

    ReDim matDados(0) As String
    
    If UseWhere Then
        NovaSQL = mvarSQL & " " & ClausulaWhere
    Else
        NovaSQL = mvarSQL
    End If
    
    If NovaSQL = "" Then
        RaiseEvent LeituraErro(-1, "Nenhum SELECT foi especificado.")
        Exit Sub
    ElseIf UCase(left(NovaSQL, 6)) <> "SELECT" Then
        RaiseEvent LeituraErro(-2, "Erro de sintaxe na cláusula SELECT.")
        Exit Sub
    End If
    
    On Error GoTo LeituraErro
    
    Maximo = UBound(matCampos)
    
    Set Registro = Conexao.OpenResultset(NovaSQL, rdOpenStatic, Options:=rdExecDirect)
    
    If Not Registro.EOF Then
        Cancelou = False
        RaiseEvent RegistroEncontrado(Registro, Cancelou)
        If Not Cancelou Then
            For Indice = 0 To Maximo Step 1
                ReDim Preserve matDados(Indice) As String
                If Not IsNull(Registro(matCampos(Indice))) Then
                    matDados(Indice) = Registro(matCampos(Indice))
                Else
                    matDados(Indice) = ""
                End If
            Next Indice
        Else
            Exit Sub
        End If
    Else
        Registro.Close
        RaiseEvent RegistroNaoExiste
        Exit Sub
    End If
    
    Registro.Close
    RaiseEvent LeituraOK
    
    Exit Sub
    
LeituraErro:

    RaiseEvent LeituraErro(Err.Number, Err.Description)
    Set Registro = Nothing
    Err.Clear
    
End Sub

Public Property Let Campos(ByVal vData As String)
    
    mvarCampos = vData
    MontarMatriz vData, matCampos

End Property

Public Property Get Campos() As String
    
    Campos = mvarCampos

End Property

Public Property Let Controles(ByVal vData As String)
    
    Dim Tamanho As Long
    Dim Indice As Long
    
    mvarControles = vData
    
    MontarMatriz vData, matControles
    
    If ControlFields Then
        Tamanho = UBound(matControles)
        ReDim matCampos(Tamanho) As String
        mvarCampos = ""
        For Indice = 0 To Tamanho Step 1
            matCampos(Indice) = Mid(matControles(Indice), 4)
            mvarCampos = mvarCampos & Mid(matControles(Indice), 4) & ", "
        Next Indice
        mvarCampos = left(mvarCampos, Len(mvarCampos) - 2)
    End If
    
End Property

Private Sub MontarMatriz(ByVal vData As String, ByRef Matriz As Variant)

    Dim Tamanho As Long
    Dim CharPos As Long
    Dim NomeLido As String
    Dim Indice As Long
    Dim PosAnt As Long
    
    Tamanho = Len(vData)
    
    ReDim Matriz(0) As String
    
    If Tamanho > 0 Then
        CharPos = InStr(vData, ";")
        If CharPos = 0 Then
            Matriz(0) = vData
        Else
            PosAnt = 1
            Indice = 0
            Do While CharPos <> 0
                NomeLido = Mid(vData, PosAnt, CharPos - PosAnt)
                PosAnt = CharPos + 2
                CharPos = InStr(CharPos + 1, vData, ";")
                Matriz(Indice) = NomeLido
                Indice = Indice + 1
                ReDim Preserve Matriz(Indice) As String
            Loop
            NomeLido = Mid(vData, PosAnt)
            Matriz(Indice) = NomeLido
        End If
    End If

End Sub

Public Property Get Controles() As String
    
    Controles = mvarControles

End Property

Public Property Let sql(ByVal vData As String)
    
    Dim CamposLidos As String
    Dim SQLSecundaria As String
    Dim Posicao1 As Long
    Dim Posicao2 As Long
    Dim TempPos As Long
    Dim ref1 As Long
    Dim ref2 As Long
    
    On Error GoTo AtribuiValor
    
    SQLSecundaria = ""
    CamposLidos = ""
    
    Posicao1 = InStr(1, vData, "[Campos]", vbTextCompare)
    If Posicao1 <> 0 Then
        CamposLidos = LerCampos(matCampos)
    End If
    
    Posicao2 = InStr(1, vData, "[ExtraFields]", vbTextCompare)
    
    If Posicao2 < Posicao1 And Posicao2 > 0 Then
        TempPos = Posicao2
        Posicao2 = Posicao1
        Posicao1 = TempPos
        ref1 = 13
        ref2 = 8
    Else
        If Posicao1 = 0 Then
            ref1 = 0
        Else
            ref1 = 8
        End If
        ref2 = 13
    End If
    
    If Posicao2 <> 0 Then
        If CamposLidos <> "" Then
            CamposLidos = CamposLidos & Mid(vData, Posicao1 + ref1, (Posicao2) - (Posicao1 + ref1))
        End If
        CamposLidos = CamposLidos & LerCampos(matExtraFields)
    End If
    
    If CamposLidos <> "" And (Posicao1 <> 0 Or Posicao2 <> 0) Then
        If Posicao1 = 0 Then
            Posicao1 = Posicao2
        ElseIf Posicao2 = 0 Then
            Posicao2 = Posicao1
            ref2 = ref1
        End If
        
        SQLSecundaria = left(vData, Posicao1 - 1) & " " & CamposLidos & Mid(vData, Posicao2 + ref2)
        vData = SQLSecundaria
    End If
    
AtribuiValor:
    mvarSQL = vData

End Property

Private Function LerCampos(ByRef Matriz) As String

    Dim Indice As Long
    Dim Maximo As Long
    Dim Retorno As String
    
    Maximo = UBound(Matriz)
    
    Retorno = ""
    For Indice = 0 To Maximo Step 1
        Retorno = Retorno & Matriz(Indice) & ", "
    Next Indice
    
    If Retorno <> "" Then
        Retorno = left(Retorno, Len(Retorno) - 2)
    End If
    
    LerCampos = Retorno

End Function

Public Property Get sql() As String
    
    sql = mvarSQL

End Property

Private Function LocalizaItemData(ByVal Indice As Long) As Long

    Dim Ponteiro As Long
    Dim Maximo As Long

    LocalizaItemData = -1
    
    Maximo = Formulario.Controls(matControles(Indice)).ListCount - 1
    
    For Ponteiro = 0 To Maximo Step 1
        If Formulario.Controls(matControles(Indice)).ItemData(Ponteiro) = Val(matDados(Indice)) Then
            LocalizaItemData = Ponteiro
            Exit Function
        End If
    Next Ponteiro

End Function

Private Function LocalizaValor(ByVal Indice As Long) As Long

    Dim Ponteiro As Long
    Dim Maximo As Long

    LocalizaValor = -1
    
    Maximo = Formulario.Controls(matControles(Indice)).ListCount - 1
    
    For Ponteiro = 0 To Maximo Step 1
        If Val(Formulario.Controls(matControles(Indice)).List(Ponteiro)) = Val(matDados(Indice)) Then
            LocalizaValor = Ponteiro
            Exit Function
        End If
    Next Ponteiro
    
End Function

Private Function LocalizaTotalText(ByVal Indice As Long) As Long

    Dim Ponteiro As Long
    Dim Maximo As Long

    LocalizaTotalText = -1
    
    Maximo = Formulario.Controls(matControles(Indice)).ListCount - 1
    
    For Ponteiro = 0 To Maximo Step 1
        If UCase(Formulario.Controls(matControles(Indice)).List(Ponteiro)) = UCase(matDados(Indice)) Then
            LocalizaTotalText = Ponteiro
            Exit Function
        End If
    Next Ponteiro
    
End Function

Private Function LocalizaRightText(ByVal Indice As Long) As Long
    
    Dim Ponteiro As Long
    Dim Maximo As Long

    On Error Resume Next

    Err.Clear

    LocalizaRightText = -1
    
    Maximo = Formulario.Controls(matControles(Indice)).ListCount - 1
    
    For Ponteiro = 0 To Maximo Step 1
        If Trim(UCase(Mid(Formulario.Controls(matControles(Indice)).List(Ponteiro), InStr(Formulario.Controls(matControles(Indice)).List(Ponteiro), "-") + 1))) = UCase(matDados(Indice)) Then
            If Err.Number = 0 Then
                LocalizaRightText = Ponteiro
                Exit Function
            End If
        End If
    Next Ponteiro
    
    Err.Clear
    
End Function

Private Function LocalizaLeftText(ByVal Indice As Long) As Long
    
    Dim Ponteiro As Long
    Dim Maximo As Long

    On Error Resume Next

    Err.Clear

    LocalizaLeftText = -1
    
    Maximo = Formulario.Controls(matControles(Indice)).ListCount - 1
    
    For Ponteiro = 0 To Maximo Step 1
        If Trim(UCase(left(Formulario.Controls(matControles(Indice)).List(Ponteiro), InStr(Formulario.Controls(matControles(Indice)).List(Ponteiro), "-") - 1))) = UCase(matDados(Indice)) Then
            If Err.Number = 0 Then
                LocalizaLeftText = Ponteiro
                Exit Function
            End If
        End If
    Next Ponteiro
    
    Err.Clear
    
End Function

Private Sub Class_Initialize()

    ReDim matControles(0) As String
    ReDim matCampos(0) As String
    ReDim matDados(0) As String

    UseWhere = False
    DisparaEventos = True
    ExecProcedure = False
    mvarSQL = ""
    mvarControles = ""
    ControlFields = True
    Adicionar = False
    mvarCampos = ""
    ClausulaWhere = ""
    IncludeExtraFields = False
    mvarExtraFields = ""
    mvarExtraValues = ""
    mvarTipoFormato = ""
    mvarNomeFormulario = ""

End Sub

Private Function Numeros2(ByVal Texto As String) As String

    Dim Maximo As Integer
    Dim Char As Integer
    Dim Charlido As String * 1
    Dim Retorno As String
    
    Maximo = Len(Texto)
    
    Retorno = ""
    For Char = 1 To Maximo Step 1
        Charlido = Mid(Texto, Char, 1)
        If IsNumeric(Charlido) Then
            Retorno = Retorno & Charlido
        End If
    Next Char
    
    Texto = Retorno
    
    Numeros2 = Texto

End Function

