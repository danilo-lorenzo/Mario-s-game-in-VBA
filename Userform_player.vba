' All comments are written in Portuguese, Please read the Readme file to understand better what each line is doing 
' Módulo teórico - Curso VBA - Jogo do Mário
' Userform Jogador
' https://www.linkedin.com/in/dltc/
'----------------------------------------------------

Option Explicit

 ' Função que cadastra o nome do jogador e o tempo inical na planilha

Private Sub cadastrar_Click()

Dim wb As Workbook: Set wb = ThisWorkbook  ' Define o livro de pastas
Dim ws As Worksheet                        ' Define a pasta de trabalho
Set ws = wb.Sheets("Mario")                ' Define a planilha

ws.Range("D9") = nome.Value                ' Escreve o nome do jogador na planilha
iniciar_contador                           ' Escreve o tempo inicial na planilha

Userform_jogo.Show                         ' Ativa Userfom do Jogo
Unload Me                                  ' Fecha Userform do jogador (esse próprio userform)



End Sub

Private Sub nome_Change()

End Sub

 ' Função que limpa as células da planilha quando o userform começa

Private Sub UserForm_Initialize()

Range("D9").Value = ""
Range("D10").Value = ""
Range("D11").Value = ""
Range("D12").Value = ""
Range("D13").Value = ""

End Sub
