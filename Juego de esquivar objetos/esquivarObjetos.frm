VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Juego de esquivar objetos"
   ClientHeight    =   4935
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   4830
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Puntuación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4815
      Begin VB.Label recordScoreLabel 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "RECORD:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label currentScoreLabel 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "ACTUAL:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   4200
      Width           =   4815
      Begin VB.CommandButton exitButton 
         Caption         =   "SALIR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   8
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton playButton 
         Caption         =   "JUGAR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Timer globalTimer 
      Interval        =   1
      Left            =   6840
      Top             =   4680
   End
   Begin VB.Timer fallingFiguresTimer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   6360
      Top             =   4680
   End
   Begin VB.PictureBox mainFigure 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   2280
      Picture         =   "esquivarObjetos.frx":0000
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   0
      Top             =   3840
      Width           =   300
   End
   Begin VB.PictureBox fallingFigure 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   2
      Left            =   4080
      Picture         =   "esquivarObjetos.frx":041B
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   5
      Top             =   0
      Width           =   300
   End
   Begin VB.PictureBox fallingFigure 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   3
      Left            =   2880
      Picture         =   "esquivarObjetos.frx":07D1
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   6
      Top             =   0
      Width           =   300
   End
   Begin VB.PictureBox fallingFigure 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   1
      Left            =   1680
      Picture         =   "esquivarObjetos.frx":0B87
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   4
      Top             =   0
      Width           =   300
   End
   Begin VB.PictureBox fallingFigure 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   600
      Picture         =   "esquivarObjetos.frx":0F3D
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   1
      Top             =   0
      Width           =   300
   End
   Begin VB.Label gameEndLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      Caption         =   "Fin del juego"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   600
      TabIndex        =   13
      Top             =   2040
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  'Opaque
      FillStyle       =   7  'Diagonal Cross
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   4815
   End
   Begin VB.Shape Shape3 
      FillStyle       =   7  'Diagonal Cross
      Height          =   495
      Left            =   0
      Top             =   4200
      Width           =   4815
   End
   Begin VB.Shape Shape2 
      FillStyle       =   7  'Diagonal Cross
      Height          =   5415
      Left            =   4440
      Top             =   -720
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillStyle       =   7  'Diagonal Cross
      Height          =   5415
      Left            =   0
      Top             =   -720
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mainFigureSpeed As Integer          'Velocidad de la figura principal
Dim fallingFigurePosition(4) As Integer 'Posiciones iniciales de las figuras que caen
Dim score As Integer                    'Almacena el puntaje actual
Dim recordScore As Integer              'Almacena el puntaje mas alto

'Finaliza el programa
Private Sub exitButton_Click()
    End
End Sub

'Controla el movimiento de la figura principal de derecha a izquierda
Private Sub mainFigure_KeyDown(KeyCode As Integer, Shift As Integer)
    'Tecla de direccion izquierda
    If (KeyCode = vbKeyLeft) Then
        mainFigure.Left = mainFigure.Left - mainFigureSpeed
    End If
    'Tecla de direccion derecha
    If (KeyCode = vbKeyRight) Then
        mainFigure.Left = mainFigure.Left + mainFigureSpeed
    End If
End Sub
'Carga inicial del formulario
Private Sub Form_Load()
    recordScore = 0
    mainFigureSpeed = 200
End Sub
'Maneja la colision de los objetos
Private Sub handleColision(i As Integer)
    'Es true si el objeto principal colisiona con cualquier objeto que cae
    If ((mainFigure.Left < fallingFigure(i).Left + fallingFigure(i).Width) And (fallingFigure(i).Left < mainFigure.Left + mainFigure.Width) And (mainFigure.Top < fallingFigure(i).Top + fallingFigure(i).Height) And (fallingFigure(i).Top < mainFigure.Top + mainFigure.Height)) Then
        'Comprueba si el puntaje actual es mayor que el puntaje mas alto obtenido
        If score > recordScore Then
            
            recordScore = score
            recordScoreLabel.Caption = recordScore 'Actualiza el label del puntaje mas alto
            MsgBox "Felicitaciones, superaste la puntuación mas alta" 'Muestra un Msgbox
        
        End If
        enableOrDisable (False) 'Desabilita o habilita elementos del formulario
        playButton.SetFocus

    End If
    
End Sub

'Evita que la figura principal se salga del marco definido
Private Sub globalTimer_Timer()

    If (mainFigure.Left <= 360) Then
        mainFigure.Left = 360
    End If
    
    If (mainFigure.Left >= 4200) Then
        mainFigure.Left = 4200
    End If
    
    Dim i As Integer
    
    
    For i = 0 To fallingFigure.Count - 1
        handleColision (i)
    Next
    
    currentScoreLabel.Caption = score 'Muesta el valor del puntaje actual
    
End Sub

Private Sub playButton_Click()
    score = 0                             'Setea el puntaje a 0
    setFallingFiguresPositions            'Figuras que caen en su posicion inicial
    mainFigure.Left = 2280                'Posicion inicial de la figura principal
    playButton.Caption = "Volver a jugar" 'Cambia el texto del boton play
    enableOrDisable (True)                'Habilita o desabilita elementos del formulario

    mainFigure.SetFocus                   'Pone el focus en la figura principal
End Sub

'Temporizador que controla el movimiento de caida de las figuras
Private Sub fallingFiguresTimer_Timer()
    Dim i As Integer
    
    For i = 0 To fallingFigure.Count - 1
    
        fallingFigure(i).Top = fallingFigure(i).Top + 50
    
        If fallingFigure(i).Top >= 4320 Then 'Si la figura sobrepasa el limite establecido
            fallingFigure(i).Top = -20       'Envia a la figura a la parte superior
            fallingFigure(i).Left = getRandomNumber(360, 4080) 'Coloca al la figura en una posicion aleatoria horizontalmente
            score = score + 1 'Se añade un punto + 1
        End If
    Next
    
End Sub
'Retorna un numero aleatoreo
Public Function getRandomNumber(min As Integer, Optional max As Integer) As Integer
    
    getRandomNumber = (CLng(min - max) * Rnd + max)
    
End Function
'Setea las posiciones iniciales de las figuras que caen
Private Sub setFallingFiguresPositions()

    'Posiciones iniciales aleatorias para el eje vertical
    fallingFigurePosition(0) = getRandomNumber(-2500)
    fallingFigurePosition(1) = getRandomNumber(-2000)
    fallingFigurePosition(2) = getRandomNumber(-4500)
    fallingFigurePosition(3) = getRandomNumber(-7500)

    Dim i As Integer
    
    For i = 0 To fallingFigure.Count - 1
        fallingFigure(i).Top = fallingFigurePosition(i)    'Coloca las figuras en sus posiciones iniciales
        fallingFigure(i).Left = getRandomNumber(360, 4080) 'Posiciones iniciales aleatorias para el eje horizontal dentro de un rango determinado
    Next


End Sub
'Habilita o desabilita componentes del formulario
Private Sub enableOrDisable(state As Boolean)
    fallingFiguresTimer.Enabled = state 'Temporizador que controla caida de las figuras
    globalTimer.Enabled = state         'Temporizador que maneja las colisiones y limites de desplazamiento de la figura principal
    playButton.Enabled = Not state      'Boton de inicio del juego
    mainFigure.Enabled = state          'Estado de la figura principal
    gameEndLabel.Visible = Not state    'Label que se muesta al finalizar el juego

End Sub



