VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "JUEGO DEL AHORCADO"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7785
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   7785
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "INFO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   20
      Top             =   1440
      Width           =   7335
      Begin VB.CommandButton playAgain 
         Caption         =   "VOLVER A JUGAR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Visible         =   0   'False
         Width           =   7095
      End
      Begin VB.Label gameInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "DFGFDGFD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   7095
      End
   End
   Begin VB.Frame wordInputFrame 
      BackColor       =   &H00E0E0E0&
      Caption         =   "INGRESO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   17
      Top             =   360
      Width           =   7335
      Begin VB.TextBox wordTextBox 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   5775
      End
      Begin VB.CommandButton play 
         BackColor       =   &H00FFFFFF&
         Caption         =   "JUGAR"
         BeginProperty Font 
            Name            =   "Segoe UI Symbol"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton close 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SALIR"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6360
      Width           =   7335
   End
   Begin VB.Timer gameTimer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   8040
      Top             =   6360
   End
   Begin VB.Frame wordViewFrame 
      BackColor       =   &H00E0E0E0&
      Caption         =   "PALABRA"
      BeginProperty Font 
         Name            =   "Segoe UI Symbol"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   7335
      Begin VB.TextBox letterContainer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   18
         Left            =   6600
         TabIndex        =   27
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox letterContainer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   17
         Left            =   6240
         TabIndex        =   26
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox letterContainer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   16
         Left            =   5880
         TabIndex        =   25
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox letterContainer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   15
         Left            =   5520
         TabIndex        =   24
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox letterContainer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   14
         Left            =   5160
         TabIndex        =   23
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox letterContainer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   13
         Left            =   4800
         TabIndex        =   15
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox letterContainer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   12
         Left            =   4440
         TabIndex        =   14
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox letterContainer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   11
         Left            =   4080
         TabIndex        =   13
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox letterContainer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   10
         Left            =   3720
         TabIndex        =   12
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox letterContainer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   9
         Left            =   3360
         TabIndex        =   11
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox letterContainer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   8
         Left            =   3000
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox letterContainer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox letterContainer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   1
         Left            =   480
         TabIndex        =   8
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox letterContainer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   2
         Left            =   840
         TabIndex        =   7
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox letterContainer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   3
         Left            =   1200
         TabIndex        =   6
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox letterContainer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   4
         Left            =   1560
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox letterContainer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   5
         Left            =   1920
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox letterContainer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   6
         Left            =   2280
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox letterContainer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   7
         Left            =   2640
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ERRORES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   240
      TabIndex        =   0
      Top             =   2280
      Width           =   7335
      Begin VB.Line rightLeg 
         BorderWidth     =   10
         Visible         =   0   'False
         X1              =   4320
         X2              =   4680
         Y1              =   2520
         Y2              =   3480
      End
      Begin VB.Line leftLeg 
         BorderWidth     =   10
         Visible         =   0   'False
         X1              =   4080
         X2              =   3720
         Y1              =   2400
         Y2              =   3480
      End
      Begin VB.Line rightArm 
         BorderWidth     =   10
         Visible         =   0   'False
         X1              =   4560
         X2              =   4920
         Y1              =   1920
         Y2              =   2520
      End
      Begin VB.Line leftArm 
         BorderWidth     =   10
         Visible         =   0   'False
         X1              =   3840
         X2              =   3360
         Y1              =   1920
         Y2              =   2520
      End
      Begin VB.Shape chest 
         FillStyle       =   0  'Solid
         Height          =   975
         Left            =   3840
         Shape           =   4  'Rounded Rectangle
         Top             =   1800
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Shape head 
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   3720
         Shape           =   3  'Circle
         Top             =   960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Line Line1 
         BorderWidth     =   15
         X1              =   2160
         X2              =   1200
         Y1              =   720
         Y2              =   1560
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   1080
         Top             =   480
         Width           =   3255
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   3015
         Left            =   1080
         Top             =   480
         Width           =   255
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   120
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Shape rope 
         BackColor       =   &H80000001&
         BorderColor     =   &H0080C0FF&
         FillColor       =   &H0080C0FF&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   4200
         Top             =   600
         Width           =   105
      End
   End
   Begin VB.Shape Shape5 
      Height          =   6855
      Left            =   120
      Top             =   240
      Width           =   7575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim wordLength As Integer
Dim i As Integer
Dim isLetterInWord As Boolean
Dim word As String
Dim letterInput As String
Dim letterInWord As String
Dim successes As Integer
Dim errors As Integer

Function verifyKey(Tecla_Presionada)
    
    Dim allowedKeys As String
    
    allowedKeys = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZabcdefghijklmnñopqrstuvwxyz" & Chr(vbKeyBack)
    
    If InStr(1, allowedKeys, Chr(Tecla_Presionada)) Then
        verifyKey = Tecla_Presionada
    Else
        verifyKey = 0
    End If
    
End Function

Private Sub close_Click()
    End
End Sub

Private Sub Form_Load()
    
    showMessage ("Ingrese una palabra y presione Jugar")
    
End Sub

Private Sub gameTimer_Timer()
    
    letterInput = InputBox("Ingrese una letra.", "Ingreso", "", 0, 0)
    
    isLetterInWord = False
    
    If letterInput <> "" Then
    
       
        
        For i = 1 To wordLength
            
            letterInWord = Mid(word, i, 1)
            
            If UCase(letterInput) = UCase(letterInWord) Then
                
                isLetterInWord = True
                showMessage ("Letra correcta: " & letterInput)
                
                Select Case i
                    
                    Case i
                        letterContainer(i - 1).Text = UCase(letterInput)
                End Select
                   
            End If
            
        Next
        
        If isLetterInWord = False Then
            
            errors = errors + 1
            showMessage ("Letra incorrecta: " & letterInput)
            If Len(letterInput) > 1 Then
                checkWord (letterInput)
            End If
            
            Select Case errors
                
                Case 1
                    rope.Visible = True
                    head.Visible = True
                Case 2
                    chest.Visible = True
                Case 3
                    leftArm.Visible = True
                Case 4
                    rightArm.Visible = True
                Case 5
                    leftLeg.Visible = True
                Case 6
                    rightLeg.Visible = True
                    
            End Select
            
        End If
        
    Else
        showMessage ("Error: Ingrese una letra")
        
    End If
    
    successes = 0
    
    For i = 0 To wordLength - 1
        
        If letterContainer(i) <> "" Then
            
            successes = successes + 1
        End If
    Next
    
    If successes = wordLength Then
        
        gameTimer.Enabled = False
        MsgBox "Felicidades, ganaste!!!"
        showMessage ("Ganador!!")
        word = ""
        playAgain.Visible = True
        
    End If
    
    If errors = 6 Then
        gameTimer.Enabled = False
        MsgBox "Perdiste"
        showMessage ("La palabra era: " & word)
        word = ""
        playAgain.Visible = True
    End If
    
End Sub

Private Sub wordTextBox_KeyPress(KeyAscii As Integer)
    KeyAscii = verifyKey(KeyAscii)
    
End Sub

Private Sub play_Click()
    
    For i = 0 To letterContainer.Count - 1
        letterContainer(i).Text = ""
    Next i
    
    word = wordTextBox.Text
    wordLength = Len(word)
    
    If (word = "") Then
        
        showMessage ("Error: Ingrese una palabra.")
        
    Else
        
        If (wordLength < 4) Then
            showMessage ("Error: Palabra muy corta")
        ElseIf wordLength > letterContainer.Count - 1 Then
            showMessage ("Error: Superaste el numero de letras permitidas")
            wordLength = 0
        Else
            
            wordInputFrame.Visible = False
            wordViewFrame.Visible = True
            showMessage ("Comienza el juego!")
            gameTimer.Enabled = True
            
            For i = 0 To wordLength - 1
                letterContainer(i).Visible = True
                
            Next i
            
        End If
    End If
End Sub

Private Sub playAgain_Click()
    
    wordTextBox.Text = ""
    wordInputFrame.Visible = True
    playAgain.Visible = False
    showMessage ("Ingrese una palabra y presione Jugar")
    errors = 0
    wordViewFrame.Visible = False
    rope.Visible = False
    head.Visible = False
    chest.Visible = False
    leftArm.Visible = False
    rightArm.Visible = False
    leftLeg.Visible = False
    rightLeg.Visible = False
    
End Sub

Private Function checkWord(letterInput As String)
    If UCase(word) = UCase(letterInput) Then
           
        Dim i As Integer
        Dim letterInWord As String
    
        For i = 1 To Len(letterInput)
            letterInWord = Mid(letterInput, i, 1)
            letterContainer(i - 1).Text = UCase(letterInWord)
        Next i
        
    Else
          showMessage ("Palabra incorrecta")
        
    End If
End Function

Private Function showMessage(message As String)
    gameInfo.Caption = message
End Function
