Dim squareValue, squareArea, squarePerimeter

squareValue = InputBox("Digite o valor do lado do quadrado")

if IsEmpty(squareValue) Then
    WScript.Quit
End if

if IsNumeric(squareValue) Then
    squareArea = squareValue * squareValue
    squarePerimeter = squareValue * 4
    MsgBox "Area do quadrado: " & squareArea & vbNewLine & _
           "Perimetro do quadrado: " & squarePerimeter, vbInformation, "Resultado"
Else
   MsgBox "Valor invalido! Por favor, digite um numero.", vbExclamation, "Erro"
end if
