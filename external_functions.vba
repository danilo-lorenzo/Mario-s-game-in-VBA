' Módulo teórico - Curso VBA - Jogo do Mário
' Adicionando funções externas
' https://www.linkedin.com/in/dltc/
'----------------------------------------------------

 ' Chama o userform jogador - Esta macro é adicionada ao botão "jogar agora" da planilha
 
Sub chamar_userform()

Userform_jogador.Show

End Sub


 ' Função que faz com que a moeda apareça e faça barulho
 
Sub pular_moeda()

 ' Condição: se o Mario pular na posição de x= 316 (exatamente em baixo da caixa) aparecer moeda

 If X = 316 Then
 
 ' Faz com que a moeda apareça
 
 Userform_jogo.moeda.Visible = True
 
 ' Chama função de toque da moeda - presente no módulo "tocar_som"
 
 tocar_moeda
 
 End If
 
 
End Sub
