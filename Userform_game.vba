' Módulo teórico - Curso VBA - Jogo do Mário
' Userform jogo
' https://www.linkedin.com/in/dltc/
'----------------------------------------------------

Option Explicit

 'Função que marca o tempo final, o tempo jogado e encerra o Jogo
    
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
terminar_contador    ' Marcação do tempo final
tempo_jogado         ' Tempo jogado = Tempo final - tempo incial

End Sub

' Define as condições de início do jogo

Private Sub UserForm_Initialize() 'Inicio da função

' Inicializando imagem do mário Mario
Userform_jogo.mario.Picture = LoadPicture(ActiveWorkbook.Path & "\frente.bmp")
' Definindo as posições inicias do mario
Userform_jogo.mario.Top = 174  ' Posição 174 no eixo Y
Userform_jogo.mario.Left = 36  ' Posição 36 no eixo X

Userform_jogo.moeda.Visible = False   ' Faz com que a moeda comece de forma invisível

End Sub 'fim da função

' Userform responsável pela movimentação do personagem

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)


' Mover para direita

    If KeyCode = 39 Then
' Movimentar mario para a direita
         mario.Left = mario.Left + 40
' Carregar foto do mario indo para a direita
         Userform_jogo.mario.Picture = LoadPicture(ActiveWorkbook.Path & "\lado_d.bmp")
' delay de 0,58 segundos
         Call DelayMs(580)
  
  
' Mover para esquerda

    ElseIf KeyCode = 37 Then
' Movimentar mario para a esquerda
       mario.Left = mario.Left - 40
' Carregar foto do mario indo para a esquerda
       Userform_jogo.mario.Picture = LoadPicture(ActiveWorkbook.Path & "\lado_e.bmp")
' delay de 0,58 segundos
       Call DelayMs(580)

' Mover para cima
    ElseIf KeyCode = 38 Then
' Movimentar mario para cima
       mario.Top = mario.Top - 40
' Carregar foto do mario pulando
       Userform_jogo.mario.Picture = LoadPicture(ActiveWorkbook.Path & "\pular.bmp")
' Tocar som do Mário pulando
       tocar_pulo
' delay de 0,58 segundos
       Call DelayMs(580)
' Fazer com que o Mário desça ao chão
       mario.Top = mario.Top + 40
' Verificar se o  mário está embaixo da caixa da moeda
 Dim X As Integer

 X = Userform_jogo.mario.Left

 If X = 316 Then                  ' Posicionar mário embaixo da caixa
 Userform_jogo.moeda.Visible = True  ' Tornar a moeda visível
 tocar_moeda                      ' Tocar som da moeda
 Range("D10") = Range("D10") + 1  ' Adicionar incremento na planilha (número de moedas)
 End If
 
' delay de 0,58 segundos
       Call DelayMs(580)
' Tornar a moeda invisível
      Userform_jogo.moeda.Visible = False


' Mover para baixo
    ElseIf KeyCode = 40 Then
' Carregar foto do mario abaixando
       Userform_jogo.mario.Picture = LoadPicture(ActiveWorkbook.Path & "\baixo.bmp")
' delay de 0,58 segundos
       Call DelayMs(580)
 End If

    End Sub



