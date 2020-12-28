' All comments are written in Portuguese, Please read the Readme file to understand better what each line is doing 
' Módulo teórico - Curso VBA - Jogo do Mário
' Funções de Contagem de tempo
' https://www.linkedin.com/in/dltc/
'----------------------------------------------------
 
 ' Torna a declaração das variáveis obrigatória
Option Explicit


 ' Função delay - usada para fazer a transição entre uma imagem do Mário e outra

 Sub DelayMs(ms As Long)

 ' Realizar espera   Application.Wait (Hora de Agora + (valor definido pelo usuário  * fator de correção))
    
    Application.Wait (Now + (ms * 0.00000001))
 
 ' Ao fim de cada transição, sempre carregamos o mário voltado para frente
    
    Userform_jogo.mario.Picture = LoadPicture(ActiveWorkbook.Path & "\frente.bmp")
        
End Sub

' Contar tempo Inical

Sub iniciar_contador()

Dim tempo_inicial As Integer
tempo_inicial = TimeValue(Now) 'Introduz hora de inicial na variável
Range("D11") = TimeValue(Now)  'Introduz hora de inicial na célula


End Sub

' Contar tempo Final

Sub terminar_contador()

Dim tempo_final As Integer
tempo_final = TimeValue(Now)   'Introduz hora final na variável
Range("D12") = TimeValue(Now)  'Introduz hora final na célula


End Sub

' Subtrair tempo inicial do final

Sub tempo_jogado()
Range("D13") = Range("D12") - Range("D11")

End Sub

