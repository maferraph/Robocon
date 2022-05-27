Attribute VB_Name = "Modulo_FuncoesGerais"
'*************************** IMPORTANTE *********************************
'Ao mudar o �ngulo do eixo "n", os valores de posi��o que ser�o alterados
'ser�o sempre do eixo "n+1"
'************************************************************************

'****************** DEFINI��O DO GRAU DOS EIXOS *************************
'Todos os 6 eixos, o valor do grau indicado � em rela��o ao observador
'olhar de frente o eixo (como se estivesse desenhado no papel), sendo
'o 1� quadrante com o eixo horizontal em 0� e o vertical em 90�, o 2�
'quadrante com o eixo vertical em 90� e o vertical em 180�, o 3� quadrante
'com o horizontal em 180� e o vertical com 270� e o 4� quadrante com
'o eixo vertical com 270� e o horizontal com 360�
'
'Todas as posi��o X,Y,Z de todos os eixos � exibido considerando a posi��o
'do robo, em rela��o ao robo ou em rela��o ao eixo 1.
'************************************************************************

'Eixo 1
'posicao zero quando o eixo olhando de cima do robo, est� 90 graus a vista do observador
Public Const ICONST_EIXO1_PASSOS As Integer = 48
Public SGVAR_EIXO1_GRAU As Single
Public Const ICONST_EIXO1_GRAUMINIMO As Integer = 0
Public Const ICONST_EIXO1_GRAUMAXIMO As Integer = 360

'Eixo 2
'posicao zero quando o eixo esta totalmente para baixo, ou seja, paralelo em relacao a mesa
Public Const ICONST_EIXO2_PASSOS As Integer = 48
Public SGVAR_EIXO2_GRAU As Single
Public Const SGCONST_EIXO2_DISTANCIA_EIXO1_EIXO2 As Single = 100
Public SGVAR_EIXO2_POSICAO_X As Single
Public SGVAR_EIXO2_POSICAO_Y As Single
Public SGVAR_EIXO2_POSICAO_Z As Single
Public Const ICONST_EIXO2_GRAUMINIMO As Integer = 0
Public Const ICONST_EIXO2_GRAUMAXIMO As Integer = 180

'Eixo 3
Public Const ICONST_EIXO3_PASSOS As Integer = 48
Public SGVAR_EIXO3_GRAU As Single
Public Const SGCONST_EIXO3_DISTANCIA_EIXO2_EIXO3 As Single = 100
Public SGVAR_EIXO3_POSICAO_X As Single
Public SGVAR_EIXO3_POSICAO_Y As Single
Public SGVAR_EIXO3_POSICAO_Z As Single
Public Const ICONST_EIXO3_GRAUMINIMO As Integer = 210
Public Const ICONST_EIXO3_GRAUMAXIMO As Integer = 150

'Eixo 4
Public Const ICONST_EIXO4_PASSOS As Integer = 48
Public SGVAR_EIXO4_GRAU As Single
Public Const SGCONST_EIXO4_DISTANCIA_EIXO3_EIXO4 As Single = 100
Public SGVAR_EIXO4_POSICAO_X As Single
Public SGVAR_EIXO4_POSICAO_Y As Single
Public SGVAR_EIXO4_POSICAO_Z As Single
Public Const ICONST_EIXO4_GRAUMINIMO As Integer = 0
Public Const ICONST_EIXO4_GRAUMAXIMO As Integer = 360

'Eixo 5
Public Const ICONST_EIXO5_PASSOS As Integer = 48
Public SGVAR_EIXO5_GRAU As Single
Public Const SGCONST_EIXO5_DISTANCIA_EIXO4_EIXO5 As Single = 100
Public SGVAR_EIXO5_POSICAO_X As Single
Public SGVAR_EIXO5_POSICAO_Y As Single
Public SGVAR_EIXO5_POSICAO_Z As Single
Public Const ICONST_EIXO5_GRAUMINIMO As Integer = 210
Public Const ICONST_EIXO5_GRAUMAXIMO As Integer = 150

'Eixo 6
Public Const ICONST_EIXO6_PASSOS As Integer = 48
Public SGVAR_EIXO6_GRAU As Single
Public Const SGCONST_EIXO6_DISTANCIA_EIXO2_EIXO3 As Single = 100
Public SGVAR_EIXO6_POSICAO_X As Single
Public SGVAR_EIXO6_POSICAO_Y As Single
Public SGVAR_EIXO6_POSICAO_Z As Single
Public Const ICONST_EIXO6_GRAUMINIMO As Integer = 0
Public Const ICONST_EIXO6_GRAUMAXIMO As Integer = 360

'Constantes e vari�veis diversas
Public Const SCONST_FORMATO_NUMERO As String = "##0.000"




Public Function Radiano2Grau(RADIANO As Single) As Single
    Radiano2Grau = (RADIANO * 3.14159265358979) / 180
End Function
Public Function Grau2Minuto2Segundo(VALOR As Single) As String
    Dim VGRAU, VMINUTO, VSEGUNDO, TEMP As Single
    Dim GRAU, MINUTO, SEGUNDO As String
    'pega graus
    VGRAU = Fix(VALOR)
    GRAU = Str(VGRAU)
    'pega minutos
    TEMP = (VALOR - VGRAU) / (1 / 60)
    VMINUTO = Fix(TEMP)
    MINUTO = Str(VMINUTO)
    'pega segundos
    TEMP = (TEMP - VMINUTO) / (1 / 60)
    VSEGUNDO = Fix(TEMP)
    SEGUNDO = Str(VSEGUNDO)
    Grau2Minuto2Segundo = GRAU & "�" & MINUTO & "'" & SEGUNDO & Chr(34)
End Function
