' Módulo teórico - Curso VBA - Jogo do Mário
' Adicionando sons ao Jogo
' https://www.linkedin.com/in/dltc/
'----------------------------------------------------

 ' Torna a declaração das variáveis obrigatória
Option Explicit

 ' Declarando função que habilita o som
Public Declare Function sndPlaySound Lib "winmm.dll" _
Alias "sndPlaySoundA" (ByVal lpszSoundName As String, _
ByVal uFlags As Long) As Long

 ' Macro que de fato executa o som

Sub PlayWavFile(WavFileName As String, Wait As Boolean)

    If Dir(WavFileName) = "" Then Exit Sub ' Nenhum arquivo de som selecionado
    If Wait Then ' Tocar o som antes de executar o restante do código
        sndPlaySound WavFileName, 0
    Else ' Tocar som enquanto o código funciona
        sndPlaySound WavFileName, 1
    End If
    
End Sub

 ' Macro que referencia o arquivo de som do pulo
Sub tocar_pulo()

    PlayWavFile ActiveWorkbook.Path & "\smw_jump.wav", True
    
End Sub

 ' Macro que referencia o arquivo de som da moeda
 
Sub tocar_moeda()

    PlayWavFile ActiveWorkbook.Path & "\smw_coin.wav", True
  
End Sub
