VERSION 5.00
Begin VB.Form Tela_Principal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RoboCon"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   8310
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FR_Eixo6 
      Caption         =   "Eixo 6"
      Height          =   615
      Left            =   120
      TabIndex        =   32
      Top             =   3600
      Width           =   8055
      Begin VB.CommandButton BT_Eixo6_Sobe 
         Caption         =   "SOBE"
         Height          =   255
         Left            =   6360
         TabIndex        =   34
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton BT_Eixo6_Desce 
         Caption         =   "DESCE"
         Height          =   255
         Left            =   7200
         TabIndex        =   33
         Top             =   240
         Width           =   735
      End
      Begin VB.Label LB_Eixo6_PosY 
         AutoSize        =   -1  'True
         Caption         =   "Posição Y:"
         Height          =   195
         Left            =   3240
         TabIndex        =   38
         Top             =   240
         Width           =   765
      End
      Begin VB.Label LB_Eixo6_PosX 
         AutoSize        =   -1  'True
         Caption         =   "Posição X:"
         Height          =   195
         Left            =   1680
         TabIndex        =   37
         Top             =   240
         Width           =   765
      End
      Begin VB.Label LB_Eixo6_Grau 
         AutoSize        =   -1  'True
         Caption         =   "Grau:"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   390
      End
      Begin VB.Label LB_Eixo6_PosZ 
         AutoSize        =   -1  'True
         Caption         =   "Posição Z:"
         Height          =   195
         Left            =   4800
         TabIndex        =   35
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.Frame FR_Eixo5 
      Caption         =   "Eixo 5"
      Height          =   615
      Left            =   120
      TabIndex        =   25
      Top             =   2880
      Width           =   8055
      Begin VB.CommandButton BT_Eixo5_Sobe 
         Caption         =   "SOBE"
         Height          =   255
         Left            =   6360
         TabIndex        =   27
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton BT_Eixo5_Desce 
         Caption         =   "DESCE"
         Height          =   255
         Left            =   7200
         TabIndex        =   26
         Top             =   240
         Width           =   735
      End
      Begin VB.Label LB_Eixo5_PosY 
         AutoSize        =   -1  'True
         Caption         =   "Posição Y:"
         Height          =   195
         Left            =   3240
         TabIndex        =   31
         Top             =   240
         Width           =   765
      End
      Begin VB.Label LB_Eixo5_PosX 
         AutoSize        =   -1  'True
         Caption         =   "Posição X:"
         Height          =   195
         Left            =   1680
         TabIndex        =   30
         Top             =   240
         Width           =   765
      End
      Begin VB.Label LB_Eixo5_Grau 
         AutoSize        =   -1  'True
         Caption         =   "Grau:"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   390
      End
      Begin VB.Label LB_Eixo5_PosZ 
         AutoSize        =   -1  'True
         Caption         =   "Posição Z:"
         Height          =   195
         Left            =   4800
         TabIndex        =   28
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.Frame FR_Eixo1 
      Caption         =   "Eixo 1"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      Begin VB.CommandButton BT_Eixo1_Desce 
         Caption         =   "DESCE"
         Height          =   255
         Left            =   7200
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton BT_Eixo1_Sobe 
         Caption         =   "SOBE"
         Height          =   255
         Left            =   6360
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.Label LB_EIXO1_GRAU 
         AutoSize        =   -1  'True
         Caption         =   "Grau:"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   390
      End
   End
   Begin VB.Frame FR_Eixo2 
      Caption         =   "Eixo 2"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   8055
      Begin VB.CommandButton BT_Eixo2_Desce 
         Caption         =   "DESCE"
         Height          =   255
         Left            =   7200
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton BT_Eixo2_Sobe 
         Caption         =   "SOBE"
         Height          =   255
         Left            =   6360
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.Label LB_Eixo2_PosZ 
         AutoSize        =   -1  'True
         Caption         =   "Posição Z:"
         Height          =   195
         Left            =   4800
         TabIndex        =   10
         Top             =   240
         Width           =   765
      End
      Begin VB.Label LB_Eixo2_Grau 
         AutoSize        =   -1  'True
         Caption         =   "Grau:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   390
      End
      Begin VB.Label LB_Eixo2_PosX 
         AutoSize        =   -1  'True
         Caption         =   "Posição X:"
         Height          =   195
         Left            =   1680
         TabIndex        =   6
         Top             =   240
         Width           =   765
      End
      Begin VB.Label LB_Eixo2_PosY 
         AutoSize        =   -1  'True
         Caption         =   "Posição Y:"
         Height          =   195
         Left            =   3240
         TabIndex        =   5
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.Frame FR_Eixo3 
      Caption         =   "Eixo 3"
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   8055
      Begin VB.CommandButton BT_Eixo3_Sobe 
         Caption         =   "SOBE"
         Height          =   255
         Left            =   6360
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton BT_Eixo3_Desce 
         Caption         =   "DESCE"
         Height          =   255
         Left            =   7200
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
      Begin VB.Label LB_Eixo3_PosY 
         AutoSize        =   -1  'True
         Caption         =   "Posição Y:"
         Height          =   195
         Left            =   3240
         TabIndex        =   17
         Top             =   240
         Width           =   765
      End
      Begin VB.Label LB_Eixo3_PosX 
         AutoSize        =   -1  'True
         Caption         =   "Posição X:"
         Height          =   195
         Left            =   1680
         TabIndex        =   16
         Top             =   240
         Width           =   765
      End
      Begin VB.Label LB_Eixo3_Grau 
         AutoSize        =   -1  'True
         Caption         =   "Grau:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   390
      End
      Begin VB.Label LB_Eixo3_PosZ 
         AutoSize        =   -1  'True
         Caption         =   "Posição Z:"
         Height          =   195
         Left            =   4800
         TabIndex        =   14
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.Frame FR_Eixo4 
      Caption         =   "Eixo 4"
      Height          =   615
      Left            =   120
      TabIndex        =   18
      Top             =   2160
      Width           =   8055
      Begin VB.CommandButton BT_Eixo4_Desce 
         Caption         =   "DESCE"
         Height          =   255
         Left            =   7200
         TabIndex        =   20
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton BT_Eixo4_Sobe 
         Caption         =   "SOBE"
         Height          =   255
         Left            =   6360
         TabIndex        =   19
         Top             =   240
         Width           =   735
      End
      Begin VB.Label LB_Eixo4_PosZ 
         AutoSize        =   -1  'True
         Caption         =   "Posição Z:"
         Height          =   195
         Left            =   4800
         TabIndex        =   24
         Top             =   240
         Width           =   765
      End
      Begin VB.Label LB_Eixo4_Grau 
         AutoSize        =   -1  'True
         Caption         =   "Grau:"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   390
      End
      Begin VB.Label LB_Eixo4_PosX 
         AutoSize        =   -1  'True
         Caption         =   "Posição X:"
         Height          =   195
         Left            =   1680
         TabIndex        =   22
         Top             =   240
         Width           =   765
      End
      Begin VB.Label LB_Eixo4_PosY 
         AutoSize        =   -1  'True
         Caption         =   "Posição Y:"
         Height          =   195
         Left            =   3240
         TabIndex        =   21
         Top             =   240
         Width           =   765
      End
   End
End
Attribute VB_Name = "Tela_Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SGVAR_TEMP As Single
Private Sub BT_Eixo1_Desce_Click()
    SGVAR_EIXO1_GRAU = SGVAR_EIXO1_GRAU - (360 / ICONST_EIXO1_PASSOS)
    ValoresEixo1
End Sub
Private Sub BT_Eixo1_Sobe_Click()
    SGVAR_EIXO1_GRAU = SGVAR_EIXO1_GRAU + (360 / ICONST_EIXO1_PASSOS)
    ValoresEixo1
End Sub
Private Sub BT_Eixo2_Desce_Click()
    SGVAR_EIXO2_GRAU = SGVAR_EIXO2_GRAU - (360 / ICONST_EIXO2_PASSOS)
    ValoresEixo2
End Sub
Private Sub BT_Eixo2_Sobe_Click()
    SGVAR_EIXO2_GRAU = SGVAR_EIXO2_GRAU + (360 / ICONST_EIXO2_PASSOS)
    ValoresEixo2
End Sub
Private Sub BT_Eixo3_Desce_Click()
    SGVAR_EIXO3_GRAU = SGVAR_EIXO3_GRAU - (360 / ICONST_EIXO3_PASSOS)
    ValoresEixo3
End Sub
Private Sub BT_Eixo3_Sobe_Click()
    SGVAR_EIXO3_GRAU = SGVAR_EIXO3_GRAU + (360 / ICONST_EIXO3_PASSOS)
    ValoresEixo3
End Sub
Private Sub BT_Eixo4_Desce_Click()
    SGVAR_EIXO4_GRAU = SGVAR_EIXO4_GRAU - (360 / ICONST_EIXO4_PASSOS)
    ValoresEixo4
End Sub
Private Sub BT_Eixo4_Sobe_Click()
    SGVAR_EIXO4_GRAU = SGVAR_EIXO4_GRAU + (360 / ICONST_EIXO4_PASSOS)
    ValoresEixo4
End Sub
Private Sub BT_Eixo5_Desce_Click()
    SGVAR_EIXO5_GRAU = SGVAR_EIXO5_GRAU - (360 / ICONST_EIXO5_PASSOS)
    ValoresEixo5
End Sub
Private Sub BT_Eixo5_Sobe_Click()
    SGVAR_EIXO5_GRAU = SGVAR_EIXO5_GRAU + (360 / ICONST_EIXO5_PASSOS)
    ValoresEixo5
End Sub
Private Sub BT_Eixo6_Desce_Click()
    SGVAR_EIXO6_GRAU = SGVAR_EIXO6_GRAU - (360 / ICONST_EIXO6_PASSOS)
    ValoresEixo6
End Sub
Private Sub BT_Eixo6_Sobe_Click()
    SGVAR_EIXO6_GRAU = SGVAR_EIXO6_GRAU + (360 / ICONST_EIXO6_PASSOS)
    ValoresEixo6
End Sub
Private Sub Form_Load()
    'zera posição dos eixos
    SGVAR_EIXO1_GRAU = 0
    SGVAR_EIXO2_GRAU = 0
    SGVAR_EIXO3_GRAU = 0
    SGVAR_EIXO4_GRAU = 0
    SGVAR_EIXO5_GRAU = 0
    SGVAR_EIXO6_GRAU = 0
    'ajusta valores de todos eixos
    ValoresEixo1
End Sub


'**********************************************************************
'                          FUNCOES DESTE SCRIPT
'**********************************************************************
Private Sub ValoresEixo1()
    'verifica grau e habilita ou não os botoes sobe e desce
    If SGVAR_EIXO1_GRAU = ICONST_EIXO1_GRAUMINIMO Then 'minimo
        BT_Eixo1_Desce.Enabled = False
    ElseIf SGVAR_EIXO1_GRAU = ICONST_EIXO1_GRAUMAXIMO Then 'maximo
        BT_Eixo1_Sobe.Enabled = False
    Else
        BT_Eixo1_Desce.Enabled = True
        BT_Eixo1_Sobe.Enabled = True
    End If
    'define valor do grau
    If SGVAR_EIXO1_GRAU = 360 Then
        SGVAR_EIXO1_GRAU = 0
    ElseIf SGVAR_EIXO1_GRAU < 0 Then
        SGVAR_EIXO1_GRAU = 360 + SGVAR_EIXO1_GRAU
    End If
    'muda captions do eixo 1
    LB_EIXO1_GRAU.Caption = "Grau: " & Grau2Minuto2Segundo(SGVAR_EIXO1_GRAU)
    'atualiza eixo 2
    ValoresEixo2
End Sub
Private Sub ValoresEixo2()
    'verifica grau e habilita ou não os botoes sobe e desce
    If SGVAR_EIXO2_GRAU = ICONST_EIXO2_GRAUMINIMO Then 'minimo
        BT_Eixo2_Desce.Enabled = False
    ElseIf SGVAR_EIXO2_GRAU = ICONST_EIXO2_GRAUMAXIMO Then 'maximo
        BT_Eixo2_Sobe.Enabled = False
    Else
        BT_Eixo2_Desce.Enabled = True
        BT_Eixo2_Sobe.Enabled = True
    End If
    'define valor do grau
    If SGVAR_EIXO2_GRAU = 360 Then
        SGVAR_EIXO2_GRAU = 0
    ElseIf SGVAR_EIXO2_GRAU < 0 Then
        SGVAR_EIXO2_GRAU = 360 + SGVAR_EIXO2_GRAU
    End If
    'eixo X - altera com movimento do eixo 1 e do eixo 2
    SGVAR_TEMP = SGCONST_EIXO2_DISTANCIA_EIXO1_EIXO2 * Cos(Radiano2Grau(SGVAR_EIXO1_GRAU)) 'calcula Y pela posição do eixo 1
    SGVAR_EIXO2_POSICAO_X = SGVAR_TEMP * Cos(Radiano2Grau(SGVAR_EIXO2_GRAU))
    'eixo Y
    SGVAR_TEMP = SGCONST_EIXO2_DISTANCIA_EIXO1_EIXO2 * Sin(Radiano2Grau(SGVAR_EIXO1_GRAU)) 'calcula Y pela posição do eixo 1
    SGVAR_EIXO2_POSICAO_Y = SGVAR_TEMP - (SGVAR_TEMP * Sin(Radiano2Grau(SGVAR_EIXO2_GRAU)))
    'eixo Z - so em relacao ao proprio grau do eixo 2
    SGVAR_EIXO2_POSICAO_Z = SGCONST_EIXO2_DISTANCIA_EIXO1_EIXO2 * Sin(Radiano2Grau(SGVAR_EIXO2_GRAU))
    'muda captions do eixo 2
    LB_Eixo2_Grau.Caption = "Grau: " & Grau2Minuto2Segundo(SGVAR_EIXO2_GRAU)
    LB_Eixo2_PosX.Caption = "Posição X: " & Format(SGVAR_EIXO2_POSICAO_X, SCONST_FORMATO_NUMERO)
    LB_Eixo2_PosY.Caption = "Posição Y: " & Format(SGVAR_EIXO2_POSICAO_Y, SCONST_FORMATO_NUMERO)
    LB_Eixo2_PosZ.Caption = "Posição Z: " & Format(SGVAR_EIXO2_POSICAO_Z, SCONST_FORMATO_NUMERO)
    'atualiza eixo 3
    ValoresEixo3
End Sub
Private Sub ValoresEixo3()
    'verifica grau e habilita ou não os botoes sobe e desce
    If SGVAR_EIXO3_GRAU = ICONST_EIXO3_GRAUMINIMO Then 'minimo
        BT_Eixo3_Desce.Enabled = False
    ElseIf SGVAR_EIXO3_GRAU = ICONST_EIXO3_GRAUMAXIMO Then 'maximo
        BT_Eixo3_Sobe.Enabled = False
    Else
        BT_Eixo3_Desce.Enabled = True
        BT_Eixo3_Sobe.Enabled = True
    End If
    'define valor do grau
    If SGVAR_EIXO3_GRAU = 360 Then
        SGVAR_EIXO3_GRAU = 0
    ElseIf SGVAR_EIXO3_GRAU < 0 Then
        SGVAR_EIXO3_GRAU = 360 + SGVAR_EIXO3_GRAU
    End If
    'eixo X - altera com movimento do eixo 1 e do eixo 2
    SGVAR_EIXO3_POSICAO_X = ((SGCONST_EIXO2_DISTANCIA_EIXO1_EIXO2 * Cos(Radiano2Grau(SGVAR_EIXO2_GRAU))) + (SGCONST_EIXO3_DISTANCIA_EIXO2_EIXO3 * Cos(Radiano2Grau(SGVAR_EIXO2_GRAU + SGVAR_EIXO3_GRAU)))) * Cos(Radiano2Grau(SGVAR_EIXO1_GRAU))
    'eixo Y - altera com movimento do eixo 1 e do eixo 2
    SGVAR_EIXO3_POSICAO_Y = ((SGCONST_EIXO2_DISTANCIA_EIXO1_EIXO2 * Sin(Radiano2Grau(90 - SGVAR_EIXO2_GRAU))) + (SGCONST_EIXO3_DISTANCIA_EIXO2_EIXO3 * Sin(Radiano2Grau(90 - SGVAR_EIXO2_GRAU - SGVAR_EIXO3_GRAU)))) * Sin(Radiano2Grau(SGVAR_EIXO1_GRAU))
    'eixo Z - em relacao ao proprio grau do eixo 2 e 3
    SGVAR_EIXO3_POSICAO_Z = SGVAR_EIXO2_POSICAO_Z + (SGCONST_EIXO3_DISTANCIA_EIXO2_EIXO3 * Sin(Radiano2Grau(SGVAR_EIXO3_GRAU + SGVAR_EIXO2_GRAU)))
    'muda captions do eixo 3
    LB_Eixo3_Grau.Caption = "Grau: " & Grau2Minuto2Segundo(SGVAR_EIXO3_GRAU)
    LB_Eixo3_PosX.Caption = "Posição X: " & Format(SGVAR_EIXO3_POSICAO_X, SCONST_FORMATO_NUMERO)
    LB_Eixo3_PosY.Caption = "Posição Y: " & Format(SGVAR_EIXO3_POSICAO_Y, SCONST_FORMATO_NUMERO)
    LB_Eixo3_PosZ.Caption = "Posição Z: " & Format(SGVAR_EIXO3_POSICAO_Z, SCONST_FORMATO_NUMERO)
    'atualiza eixo 4
    ValoresEixo4
End Sub
Private Sub ValoresEixo4()
    'verifica grau e habilita ou não os botoes sobe e desce
    If SGVAR_EIXO4_GRAU = ICONST_EIXO4_GRAUMINIMO Then 'minimo
        BT_Eixo4_Desce.Enabled = False
    ElseIf SGVAR_EIXO4_GRAU = ICONST_EIXO4_GRAUMAXIMO Then 'maximo
        BT_Eixo4_Sobe.Enabled = False
    Else
        BT_Eixo4_Desce.Enabled = True
        BT_Eixo4_Sobe.Enabled = True
    End If
    'define valor do grau
    If SGVAR_EIXO4_GRAU = 360 Then
        SGVAR_EIXO4_GRAU = 0
    ElseIf SGVAR_EIXO4_GRAU < 0 Then
        SGVAR_EIXO4_GRAU = 360 + SGVAR_EIXO4_GRAU
    End If
    'eixo X - altera com movimento dos eixos 1,2 e 3
    SGVAR_EIXO4_POSICAO_X = ((SGCONST_EIXO2_DISTANCIA_EIXO1_EIXO2 * Cos(Radiano2Grau(SGVAR_EIXO2_GRAU))) + ((SGCONST_EIXO3_DISTANCIA_EIXO2_EIXO3 + SGCONST_EIXO4_DISTANCIA_EIXO3_EIXO4) * Cos(Radiano2Grau(SGVAR_EIXO2_GRAU + SGVAR_EIXO3_GRAU)))) * Cos(Radiano2Grau(SGVAR_EIXO1_GRAU))
    'eixo Y - altera com movimento dos eixos 1,2 e 3
    SGVAR_EIXO4_POSICAO_Y = ((SGCONST_EIXO2_DISTANCIA_EIXO1_EIXO2 * Sin(Radiano2Grau(90 - SGVAR_EIXO2_GRAU))) + ((SGCONST_EIXO3_DISTANCIA_EIXO2_EIXO3 + SGCONST_EIXO4_DISTANCIA_EIXO3_EIXO4) * Sin(Radiano2Grau(90 - SGVAR_EIXO2_GRAU - SGVAR_EIXO3_GRAU)))) * Sin(Radiano2Grau(SGVAR_EIXO1_GRAU))
    'eixo Z - em relacao ao proprio grau dos eixos 1,2 e 3
    SGVAR_EIXO4_POSICAO_Z = SGVAR_EIXO2_POSICAO_Z + ((SGCONST_EIXO3_DISTANCIA_EIXO2_EIXO3 + SGCONST_EIXO4_DISTANCIA_EIXO3_EIXO4) * Sin(Radiano2Grau(SGVAR_EIXO3_GRAU + SGVAR_EIXO2_GRAU)))
    'muda captions do eixo 4
    LB_Eixo4_Grau.Caption = "Grau: " & Grau2Minuto2Segundo(SGVAR_EIXO4_GRAU)
    LB_Eixo4_PosX.Caption = "Posição X: " & Format(SGVAR_EIXO4_POSICAO_X, SCONST_FORMATO_NUMERO)
    LB_Eixo4_PosY.Caption = "Posição Y: " & Format(SGVAR_EIXO4_POSICAO_Y, SCONST_FORMATO_NUMERO)
    LB_Eixo4_PosZ.Caption = "Posição Z: " & Format(SGVAR_EIXO4_POSICAO_Z, SCONST_FORMATO_NUMERO)
    'atualiza eixo 5
    ValoresEixo5
End Sub
Private Sub ValoresEixo5()
    'verifica grau e habilita ou não os botoes sobe e desce
    If SGVAR_EIXO5_GRAU = ICONST_EIXO5_GRAUMINIMO Then 'minimo
        BT_Eixo5_Desce.Enabled = False
    ElseIf SGVAR_EIXO5_GRAU = ICONST_EIXO5_GRAUMAXIMO Then 'maximo
        BT_Eixo5_Sobe.Enabled = False
    Else
        BT_Eixo5_Desce.Enabled = True
        BT_Eixo5_Sobe.Enabled = True
    End If
    'define valor do grau
    If SGVAR_EIXO5_GRAU = 360 Then
        SGVAR_EIXO5_GRAU = 0
    ElseIf SGVAR_EIXO5_GRAU < 0 Then
        SGVAR_EIXO5_GRAU = 360 + SGVAR_EIXO5_GRAU
    End If
    'eixo X - altera com movimento dos eixos 1,2,3 e 4
    SGVAR_EIXO5_POSICAO_X = ((SGCONST_EIXO2_DISTANCIA_EIXO1_EIXO2 * Cos(Radiano2Grau(SGVAR_EIXO2_GRAU))) + ((SGCONST_EIXO3_DISTANCIA_EIXO2_EIXO3 + SGCONST_EIXO4_DISTANCIA_EIXO3_EIXO4 + SGCONST_EIXO5_DISTANCIA_EIXO4_EIXO5) * Cos(Radiano2Grau(SGVAR_EIXO2_GRAU + SGVAR_EIXO3_GRAU + SGVAR_EIXO5_GRAU)))) * Cos(Radiano2Grau(SGVAR_EIXO1_GRAU))
    'eixo Y - altera com movimento dos eixos 1,2,3 e 4
    SGVAR_EIXO5_POSICAO_Y = ((SGCONST_EIXO2_DISTANCIA_EIXO1_EIXO2 * Sin(Radiano2Grau(90 - SGVAR_EIXO2_GRAU))) + ((SGCONST_EIXO3_DISTANCIA_EIXO2_EIXO3 + SGCONST_EIXO4_DISTANCIA_EIXO3_EIXO4 + SGCONST_EIXO5_DISTANCIA_EIXO4_EIXO5) * Sin(Radiano2Grau(90 - SGVAR_EIXO2_GRAU - SGVAR_EIXO3_GRAU - SGVAR_EIXO5_GRAU)))) * Sin(Radiano2Grau(SGVAR_EIXO1_GRAU))
    'eixo Z - em relacao ao proprio grau dos eixos 1,2,3 e 4
    SGVAR_EIXO5_POSICAO_Z = SGVAR_EIXO2_POSICAO_Z + ((SGCONST_EIXO3_DISTANCIA_EIXO2_EIXO3 + SGCONST_EIXO4_DISTANCIA_EIXO3_EIXO4 + SGCONST_EIXO5_DISTANCIA_EIXO4_EIXO5) * Sin(Radiano2Grau(SGVAR_EIXO2_GRAU + SGVAR_EIXO3_GRAU + SGVAR_EIXO5_GRAU)))
    'muda captions do eixo 5
    LB_Eixo5_Grau.Caption = "Grau: " & Grau2Minuto2Segundo(SGVAR_EIXO5_GRAU)
    LB_Eixo5_PosX.Caption = "Posição X: " & Format(SGVAR_EIXO5_POSICAO_X, SCONST_FORMATO_NUMERO)
    LB_Eixo5_PosY.Caption = "Posição Y: " & Format(SGVAR_EIXO5_POSICAO_Y, SCONST_FORMATO_NUMERO)
    LB_Eixo5_PosZ.Caption = "Posição Z: " & Format(SGVAR_EIXO5_POSICAO_Z, SCONST_FORMATO_NUMERO)
    'atualiza eixo 6
    ValoresEixo6
End Sub
Private Sub ValoresEixo6()
    'verifica grau e habilita ou não os botoes sobe e desce
    If SGVAR_EIXO6_GRAU = ICONST_EIXO6_GRAUMINIMO Then 'minimo
        BT_Eixo6_Desce.Enabled = False
    ElseIf SGVAR_EIXO6_GRAU = ICONST_EIXO6_GRAUMAXIMO Then 'maximo
        BT_Eixo6_Sobe.Enabled = False
    Else
        BT_Eixo6_Desce.Enabled = True
        BT_Eixo6_Sobe.Enabled = True
    End If
    'define valor do grau
    If SGVAR_EIXO6_GRAU = 360 Then
        SGVAR_EIXO6_GRAU = 0
    ElseIf SGVAR_EIXO6_GRAU < 0 Then
        SGVAR_EIXO6_GRAU = 360 + SGVAR_EIXO6_GRAU
    End If
    
    'muda captions do eixo 3
    LB_Eixo6_Grau.Caption = "Grau: " & Grau2Minuto2Segundo(SGVAR_EIXO6_GRAU)
    LB_Eixo6_PosX.Caption = "Posição X: " & Format(SGVAR_EIXO6_POSICAO_X, SCONST_FORMATO_NUMERO)
    LB_Eixo6_PosY.Caption = "Posição Y: " & Format(SGVAR_EIXO6_POSICAO_Y, SCONST_FORMATO_NUMERO)
    LB_Eixo6_PosZ.Caption = "Posição Z: " & Format(SGVAR_EIXO6_POSICAO_Z, SCONST_FORMATO_NUMERO)
End Sub

