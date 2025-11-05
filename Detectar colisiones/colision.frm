VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   7440
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   4755
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer global 
      Interval        =   10
      Left            =   1320
      Top             =   6480
   End
   Begin VB.Timer mov1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   840
      Top             =   6480
   End
   Begin VB.Timer mov2 
      Interval        =   10
      Left            =   360
      Top             =   6480
   End
   Begin VB.PictureBox figB1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   2400
      Picture         =   "colision.frx":0000
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   1
      Top             =   3360
      Width           =   300
   End
   Begin VB.PictureBox figG1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   2160
      Picture         =   "colision.frx":03B6
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   0
      Top             =   480
      Width           =   300
   End
   Begin VB.Shape Shape4 
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
      Top             =   6960
      Width           =   4815
   End
   Begin VB.Shape Shape2 
      FillStyle       =   7  'Diagonal Cross
      Height          =   8175
      Left            =   4440
      Top             =   -720
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillStyle       =   7  'Diagonal Cross
      Height          =   8175
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
Dim velMov As Integer
Private Sub figG1_KeyDown(KeyCode As Integer, Shift As Integer)


    If (KeyCode = vbKeyDown) Then
        figG1.Top = figG1.Top + velMov
    End If
    
    If (KeyCode = vbKeyUp) Then
        figG1.Top = figG1.Top - velMov
    End If
    
    If (KeyCode = vbKeyLeft) Then
        figG1.Left = figG1.Left - velMov
    End If
    
    If (KeyCode = vbKeyRight) Then
        figG1.Left = figG1.Left + velMov
    End If



End Sub

Private Sub Form_Load()

velMov = 100

End Sub

Private Sub colision()

    If ((figG1.Left < figB1.Left + figB1.Width) And (figB1.Left < figG1.Left + figG1.Width) And (figG1.Top < figB1.Top + figB1.Height) And (figB1.Top < figG1.Top + figG1.Height)) Then
    
        MsgBox ("Colisionaste con el objeto")
        figG1.Top = 480

    End If
    
End Sub
Private Sub global_Timer()

    If (figG1.Left <= 360) Then
        figG1.Left = 360
    End If
    
    If (figG1.Left >= 4200) Then
        figG1.Left = 4200
    End If
    
    If (figG1.Top <= 480) Then
        figG1.Top = 480
    End If
    
    If (figG1.Top >= 6720) Then
        figG1.Top = 6720
    End If
    
    colision

End Sub

Private Sub mov1_Timer()

    figB1.Left = figB1.Left + 50
    
    If figB1.Left >= 4200 Then
    
        mov1.Enabled = False
        mov2.Enabled = True
        
    End If
    
End Sub

Private Sub mov2_Timer()

    figB1.Left = figB1.Left - 50

    If figB1.Left <= 360 Then
    
       mov1.Enabled = True
       mov2.Enabled = False
       
    End If
 
End Sub

