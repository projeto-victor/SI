Dim inputValue, number, predecessor, successor

inputValue = InputBox("Digite um numero inteiro ou decimal:", "Entrada de Numero")

If IsEmpty(inputValue) Then
    WScript.Quit
End If

If IsNumeric(inputValue) Then
    predecessor = inputValue - 1
    successor = inputValue + 1
    MsgBox "Numero anterior: " & predecessor & vbNewLine & _
           "Numero posterior:  " & successor, vbInformation, "Resultado"
Else
    MsgBox "Valor invalido! Por favor, digite um numero.", vbExclamation, "Erro"
End If