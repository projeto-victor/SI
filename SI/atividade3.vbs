Dim num1, num2, num3, greater, lower

num1 = InputBox("Digite o valor do primeiro numero: ")
num2 = InputBox("Digite o valor do segundo numero: ")
num3 = InputBox("Digite o valor do terceiro numero: ")

if IsEmpty(num1) Or IsEmpty(num2) Or IsEmpty(num3) Then
    WScript.Quit
End if

if IsNumeric(num1) And IsNumeric(num2) And IsNumeric(num3) Then
   
    if num1 < num2 And num1 < num3 Then
        lower = num1
        if num2 >= num3 Then
            greater = num2
        else 
            greater = num3
        end if
    elseif num1 >= num2 Then
        if num2 >= num3 Then
            lower = num3
            greater = num1
        else
            lower = num2
            if num1 >= num3 Then
                greater = num1
            else
                greater = num3
            end if
        end if
    else
        greater = num2
        lower = num3
    end if  
     MsgBox "menor numero: " & lower & vbNewLine & _
           "maior numero:  " & greater, vbInformation, "Resultado"
Else
    MsgBox "Valor invalido! Por favor, digite um numero.", vbExclamation, "Erro"
End if