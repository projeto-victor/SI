Dim palavras(19), voz, total, ordem, acertos, pulo, i, idx, palavra, resposta, repetiu, jogarNovamente

Set voz = CreateObject("SAPI.SpVoice")

palavras(0) = "cachorro"
palavras(1) = "gato"
palavras(2) = "elefante"
palavras(3) = "computador"
palavras(4) = "programacao"
palavras(5) = "python"
palavras(6) = "algoritmo"
palavras(7) = "dados"
palavras(8) = "inteligencia"
palavras(9) = "artificial"
palavras(10) = "voz"
palavras(11) = "reconhecimento"
palavras(12) = "jogo"
palavras(13) = "aprender"
palavras(14) = "divertido"
palavras(15) = "desafio"
palavras(16) = "palavra"
palavras(17) = "texto"
palavras(18) = "sonoro"
palavras(19) = "final"

total = 20

MsgBox "JOGO DE PALAVRAS POR VOZ" & vbNewLine & _
       "Regras: ouça a palavra e digite exatamente." & vbNewLine & _
       "Atencao: todas as palavras nao possuem acentos graficos, evite o uso para progredir no jogo" & vbNewLine & _
       "Comandos: 1 = ouvir de novo (1 vez por palavra), 2 = pular (1 vez no jogo), 3 = ordem (já definida)"

Do
    ordem = InputBox("Digite A para ordem crescente ou D para decrescente")
    If ordem = "A" Then
        ordem = "crescente"
    Else
        ordem = "decrescente"
    End If
    
    acertos = 0
    pulo = 0
    
    For i = 0 To total - 1
        If ordem = "crescente" Then
            idx = i
        Else
            idx = total - 1 - i
        End If
        
        palavra = palavras(idx)
        repetiu = False
        
        Do
            If repetiu = False Then
                voz.Speak palavra
            End If
            
            resposta = InputBox("Palavra " & (acertos + 1) & " de " & total & vbNewLine & _
                                "Digite a palavra ou 1, 2, 3:", "Jogo")
            
            If resposta = "" Then WScript.Quit
            
            If resposta = "1" Then
                If repetiu = True Then
                    MsgBox "Você já ouviu esta palavra duas vezes!"
                Else
                    repetiu = True
                    voz.Speak palavra
                End If
            ElseIf resposta = "2" Then
                If pulo = 1 Then
                    MsgBox "Você já usou seu pulo!"
                Else
                    pulo = 1
                    MsgBox "Palavra pulada!"
                    Exit Do
                End If
            ElseIf resposta = "3" Then
                MsgBox "A ordem já foi escolhida no início."
            Else
                If LCase(resposta) = LCase(palavra) Then
                    acertos = acertos + 1
                    Exit Do
                Else
                    MsgBox "Errou! A palavra certa era: " & palavra
                    MsgBox "Você acertou " & acertos & " palavras."
                    Exit For
                End If
            End If
        Loop
    Next
    
    If acertos = total Then
        MsgBox "Parabéns! Você acertou todas as 20 palavras!"
    Else
        MsgBox "Fim de jogo. Acertos: " & acertos
    End If
    
    jogarNovamente = MsgBox("Deseja jogar novamente?", vbYesNo)
Loop While jogarNovamente = vbYes

MsgBox "Obrigado por jogar!"