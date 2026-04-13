Dim salario, qtdSalarios, salarioBruto, inss, salarioLiquido
Dim aliquota

salario = InputBox("Valor do salario")

qtdSalarios = InputBox("Digite a quantidade de salarios mínimos do funcionario:")

If IsEmpty(qtdSalarios) Then
    WScript.Quit
End If

If IsNumeric(qtdSalarios) And CDbl(qtdSalarios) >= 0 Then
    qtdSalarios = CDbl(qtdSalarios)
    salarioBruto = qtdSalarios * salario

    if salario <= 1621.00 Then
            aliquota = 0.075  
    elseif salario <= 2902.84 Then
            aliquota = 0.09   
    elseif salario <= 4354.27 Then
            aliquota = 0.12  
    else
        aliquota = 0.14 
    end if
    
    inss = salarioBruto * aliquota
    salarioLiquido = salarioBruto - inss

    MsgBox "Salario Bruto: R$ " & FormatNumber(salarioBruto, 2) & vbNewLine & _
           "INSS (" & FormatPercent(aliquota, 0) & "): R$ " & FormatNumber(inss, 2) & vbNewLine & _
           "Salario Liquido: R$ " & FormatNumber(salarioLiquido, 2), _
           vbInformation, "Resultado da Folha"
Else
    MsgBox "Valor invalido! Digite um número válido e positivo.", vbExclamation, "Erro"
End If