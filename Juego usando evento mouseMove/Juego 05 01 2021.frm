VERSION 5.00
Begin VB.Form Juego 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Juego en Visual Basic 6.0"
   ClientHeight    =   7830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame levelOptionsFrame 
      BackColor       =   &H00404040&
      Caption         =   "Niveles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   5520
      Width           =   10335
      Begin VB.OptionButton levelOption 
         BackColor       =   &H00404040&
         Caption         =   "Nivel 1"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton levelOption 
         BackColor       =   &H00404040&
         Caption         =   "Nivel 3"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   2280
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton levelOption 
         BackColor       =   &H00404040&
         Caption         =   "Nivel 2"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   6120
      Width           =   10335
      Begin VB.Frame Frame5 
         BackColor       =   &H00000000&
         Caption         =   "Objetivo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1335
         Left            =   4080
         TabIndex        =   9
         Top             =   120
         Width           =   2175
         Begin VB.Label targetScoreOutput 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Impact"
               Size            =   48
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   1095
            Left            =   0
            TabIndex        =   10
            Top             =   120
            Width           =   2175
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00000000&
         Caption         =   "Tiempo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1335
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   1575
         Begin VB.Label timeOutput 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Impact"
               Size            =   48
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   1095
            Left            =   120
            TabIndex        =   7
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00000000&
         Caption         =   "Puntaje"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1335
         Left            =   1800
         TabIndex        =   4
         Top             =   120
         Width           =   2175
         Begin VB.Label currentScoreOutput 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Impact"
               Size            =   48
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   1095
            Left            =   120
            TabIndex        =   5
            Top             =   120
            Width           =   1935
         End
      End
      Begin VB.CommandButton salir 
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6360
         TabIndex        =   3
         Top             =   840
         Width           =   3855
      End
      Begin VB.CommandButton playButton 
         Caption         =   "Jugar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6360
         TabIndex        =   2
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Frame gameFrame 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10335
      Begin VB.Image redFigs 
         Height          =   750
         Index           =   1
         Left            =   5280
         Picture         =   "Juego 05 01 2021.frx":0000
         Stretch         =   -1  'True
         Top             =   4080
         Width           =   750
      End
      Begin VB.Image redFigs 
         Height          =   1215
         Index           =   0
         Left            =   1680
         Picture         =   "Juego 05 01 2021.frx":07AB
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label infoLabel 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SELECCIONE EL NIVEL Y PRESIONE JUGAR PARA QUE EMPIECE EL JUEGO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   720
         TabIndex        =   8
         Top             =   120
         Width           =   9135
      End
      Begin VB.Image greenFig 
         Height          =   1170
         Left            =   4320
         Picture         =   "Juego 05 01 2021.frx":0F56
         Stretch         =   -1  'True
         Top             =   2280
         Width           =   1005
      End
   End
   Begin VB.Timer remmainingTimeTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   480
      Top             =   6840
   End
   Begin VB.Timer figuresMovementTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   6840
   End
End
Attribute VB_Name = "Juego"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim score As Integer
Dim remainingTime As Integer

Private Sub gameEnd(win As Boolean)
    score = 0
    enableOrDisable (False)
    levelOptionsFrame.Enabled = True
    remmainingTimeTimer.Enabled = False
    
    Dim i As Integer
    
    For i = 0 To levelOption.Count - 1
        levelOption(i).Value = False
    Next
    
    If win Then
        MsgBox "Felicitaciones. Completaste el nivel.", vbExclamation, "Ganaste"
    Else
        MsgBox "Perdiste. Se termino el tiempo", vbInformation, "Perdiste"
    End If
    
End Sub

Private Sub greenFig_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    score = score + 1
    currentScoreOutput.Caption = score
    If score = targetScoreOutput Then gameEnd (True)
End Sub

Private Sub redFigs_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    score = score - 1
    currentScoreOutput.Caption = score
End Sub



Private Sub playButton_Click()
    
    enableOrDisable (True)
    playButton.Enabled = False
    levelOptionsFrame.Enabled = False
    currentScoreOutput.Caption = "0"
    
End Sub

Private Sub levelOption_Click(Index As Integer)
    Dim result As Integer
    Select Case Index
    Case 0
        result = setGameConfigs(60, 500, 700, False)
    Case 1
        result = setGameConfigs(50, 700, 600, False)
    Case 2
        redFigs(1).Visible = True
        result = setGameConfigs(40, 1000, 500, True)
    End Select
End Sub

Private Function setGameConfigs(ByVal time As Integer, ByVal targetScore As Integer, ByVal timerInterval As Integer, addFigure As Boolean)
    timeOutput = time
    targetScoreOutput = targetScore
    figuresMovementTimer.Interval = timerInterval
    playButton.Enabled = True
    currentScoreOutput = 0
    remainingTime = time
    
End Function

Private Sub salir_Click()
    End
    
End Sub

Private Sub figuresMovementTimer_Timer()
    
    greenFig.Top = setRandomPositionX(100)
    greenFig.Left = setRandomPositionY(3)
    
    Dim i As Integer
    
    For i = 0 To redFigs.Count - 1
        redFigs(i).Top = setRandomPositionX(100)
        redFigs(i).Left = setRandomPositionY(6)
        
    Next
    
End Sub

Public Function setRandomPositionX(x As Integer) As Integer
    
    setRandomPositionX = Int(CLng(x - 4200) * Rnd + 4200)
    
End Function

Public Function setRandomPositionY(x As Integer) As Integer
    
    setRandomPositionY = (CLng(x - 9480) * Rnd + 9480)
    
End Function

Private Sub remmainingTimeTimer_Timer()
    
    remainingTime = remainingTime - 1
    timeOutput.Caption = remainingTime
    
    If remainingTime = 0 Then
        enableOrDisable (False)
        gameEnd (False)
        
    End If
    
End Sub

Private Sub enableOrDisable(state As Boolean)
    
    remmainingTimeTimer.Enabled = state
    figuresMovementTimer.Enabled = state
    gameFrame.Enabled = state
    infoLabel.Visible = Not state
      
End Sub
