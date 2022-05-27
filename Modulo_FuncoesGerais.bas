Attribute VB_Name = "Modulo_FuncoesGerais"
'*************************** IMPORTANTE *********************************
'Ao mudar o ângulo do eixo "n", os valores de posição que serão alterados
'serão sempre do eixo "n+1"
'************************************************************************

'****************** DEFINIÇÃO DO GRAU DOS EIXOS *************************
'Todos os 6 eixos, o valor do grau indicado é em relação ao observador
'olhar de frente o eixo (como se estivesse desenhado no papel), sendo
'o 1º quadrante com o eixo horizontal em 0º e o vertical em 90º, o 2º
'quadrante com o eixo vertical em 90º e o vertical em 180º, o 3º quadrante
'com o horizontal em 180º e o vertical com 270º e o 4º quadrante com
'o eixo vertical com 270º e o horizontal com 360º
'
'Todas as posição X,Y,Z de todos os eixos é exibido considerando a posição
'do robo, em relação ao robo ou em relação ao eixo 1.
'************************************************************************

'Eixo 1
'posicao zero quando o eixo olhando de cima do robo, está 90 graus a vista do observador
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

'Constantes e variáveis diversas
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
    Grau2Minuto2Segundo = GRAU & "º" & MINUTO & "'" & SEGUNDO & Chr(34)
End Function
