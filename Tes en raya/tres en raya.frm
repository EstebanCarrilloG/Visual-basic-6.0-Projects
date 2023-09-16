VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Tres en raya"
   ClientHeight    =   5760
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleMode       =   0  'User
   ScaleWidth      =   8520
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton validateInfo 
      Caption         =   "VALIDAR"
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
      TabIndex        =   11
      Top             =   1920
      Width           =   4335
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   4815
      Left            =   4560
      TabIndex        =   9
      ToolTipText     =   "240"
      Top             =   840
      Width           =   3855
      Begin VB.CommandButton cell 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   8
         Left            =   2520
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2640
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cell 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   7
         Left            =   1320
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2640
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cell 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   6
         Left            =   120
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2640
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cell 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   5
         Left            =   2520
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1440
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cell 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   4
         Left            =   1320
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1440
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cell 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   3
         Left            =   120
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1440
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cell 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   2
         Left            =   2520
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cell 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   1
         Left            =   1320
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cell 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   0
         Left            =   120
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton playButton 
         BackColor       =   &H0000FF00&
         Caption         =   "JUGAR"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3960
         Width           =   3615
      End
   End
   Begin VB.CommandButton close 
      Caption         =   "SALIR"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2520
      TabIndex        =   8
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "MARCADOR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2175
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   4335
      Begin VB.Label info 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Segoe UI Symbol"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   1815
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.CommandButton restartButton 
      Caption         =   "REINICIAR"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   4800
      Width           =   1935
   End
   Begin VB.TextBox playerTwoInput 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   2
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox playerOneInput 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2160
      MaxLength       =   10
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "JUGADOR 2:""X"""
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "JUGADOR 1:""O"""
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label infoLabel 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Roman"
         Size            =   21.75
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim playerOneName   As String
Dim playerTwoName   As String
Dim X               As Integer
Dim infoLabelp1     As Integer
Dim infoLabelp2     As Integer
Dim results(9, 3)   As Integer
Dim row             As Integer

Private Sub clear()
    mensaje
    gameCells ("disable")
    playButton.Enabled = True
    playButton.Caption = "VOLVER A JUGAR"
    
End Sub
Private Sub cell_Click(Index As Integer)
    If (X = 0) Then
        cell(Index).Caption = "O"
        cell(Index).Enabled = False
        X = 1
    Else
        cell(Index).Caption = "X"
        cell(Index).Enabled = False
        X = 0
    End If
    
    verify
    
End Sub

Private Sub close_Click()
End
End Sub

Private Sub validateInfo_Click()
    playerOneName = playerOneInput.Text
    playerTwoName = playerTwoInput.Text
    
    If playerOneName = "" And playerTwoName = "" Then
        infoLabel.Caption = "Ingrese los nombres de los jugadores"
        
    End If
    
    If playerOneName <> "" And playerTwoName <> "" Then
        infoLabel.Caption = "Presione jugar para continuar "
        playButton.Enabled = True
        playerOneInput.Enabled = False
        playerTwoInput.Enabled = False
    End If
    If playerOneName = "" Then
        infoLabel.Caption = "Ingrese el nombre del Jugador 1"
    End If
    
    If playerTwoName = "" Then
        infoLabel.Caption = "Ingrese el nombre del Jugador 2"
    End If
    
End Sub

Private Sub Form_Load()
    results(0, 0) = 0: results(0, 1) = 1: results(0, 2) = 2
    results(1, 0) = 3: results(1, 1) = 4: results(1, 2) = 5
    results(2, 0) = 6: results(2, 1) = 7: results(2, 2) = 8
    results(3, 0) = 0: results(3, 1) = 3: results(3, 2) = 6
    results(4, 0) = 1: results(4, 1) = 4: results(4, 2) = 7
    results(5, 0) = 2: results(5, 1) = 5: results(5, 2) = 8
    results(6, 0) = 0: results(6, 1) = 4: results(6, 2) = 8
    results(7, 0) = 2: results(7, 1) = 4: results(7, 2) = 6

    gameCells ("disable")
    randomize
    
End Sub

Private Sub verify()
    
    For row = 0 To 7
        
        If cell(results(row, 0)).Caption = "O" And cell(results(row, 1)).Caption = "O" And cell(results(row, 2)).Caption = "O" Then
            infoLabelp1 = infoLabelp1 + 1
            infoLabel.Caption = playerOneName & " ES EL GANADOR"
            clear
        ElseIf cell(results(row, 0)).Caption = "X" And cell(results(row, 1)).Caption = "X" And cell(results(row, 2)).Caption = "X" Then
            infoLabelp2 = infoLabelp2 + 1
            infoLabel.Caption = playerTwoName & " ES EL GANADOR"
            clear
        End If
        
    Next
    
    If cell(0).Caption <> "" And cell(1).Caption <> "" And cell(2).Caption <> "" And cell(3).Caption <> "" And cell(4).Caption <> "" And cell(5).Caption <> "" And cell(6).Caption <> "" And cell(7).Caption <> "" And cell(8).Caption <> "" Then
        infoLabel.Caption = "ES UN EMPATE"
        randomize
        clear
    End If
    
End Sub

Private Sub playButton_Click()
    
    playerOneInput.Enabled = False
    playerTwoInput.Enabled = False
    
    infoLabel.Caption = ""
    
    mensaje
    If X = 0 Then
        MsgBox playerOneName & " Empieza", vbInformation, "TRES EN RAYA"
    Else
        MsgBox playerTwoName & " Empieza", vbInformation, "TRES EN RAYA"
    End If
    
    gameCells ("enable")
    gameCells ("clear")
    
    playButton.Enabled = False
    
End Sub

Private Sub mensaje()
    info.Caption = playerOneName & ": " & infoLabelp1 & vbCrLf & playerTwoName & ": " & infoLabelp2
End Sub

Private Sub restartButton_Click()
    
    gameCells ("disable")
    gameCells ("clear")
    
    playButton.Enabled = True
    playerOneInput.Text = ""
    playerTwoInput.Text = ""
    playerOneInput.Enabled = True
    playerTwoInput.Enabled = True
    
    infoLabelp1 = 0
    infoLabelp2 = 0
    
    info.Caption = ""
    infoLabel.Caption = "Ingrese los nombres de los jugadores"
    playButton.Caption = "JUGAR"
    playButton.Enabled = False
    
End Sub

Private Sub playerOneInput_KeyPress(KeyAscii As Integer)
    KeyAscii = Verificar_Tecla(KeyAscii)
End Sub

Private Sub playerTwoInput_KeyPress(KeyAscii As Integer)
    KeyAscii = Verificar_Tecla(KeyAscii)
End Sub

Function Verificar_Tecla(Tecla_Presionada)
    
    Dim Teclas      As String
    
    Teclas = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZabcdefghijklmnñopqrstuvwxyz" & Chr(vbKeyBack)
    
    If InStr(1, Teclas, Chr(Tecla_Presionada)) Then
        
        Verificar_Tecla = Tecla_Presionada
    Else
        
        Verificar_Tecla = 0
    End If
    
End Function

Function gameCells(condition)
    
    Dim i As Integer
    
    If (condition = "disable") Then
        For i = 0 To cell.Count - 1
            cell(i).Enabled = False
        Next
    ElseIf (condition = "enable") Then
        For i = 0 To cell.Count - 1
            cell(i).Enabled = True
        Next
    ElseIf (condition = "clear") Then
        For i = 0 To cell.Count - 1
            cell(i).Caption = ""
        Next
    End If
    
End Function

Private Sub playBtnTimer_Timer()
    playButton.Enabled = False
End Sub

Function randomize()
    X = (CLng(0 - 1) * Rnd + 1)
End Function
