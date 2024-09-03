VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Juego de esquivar objetos"
   ClientHeight    =   4950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   4845
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame info 
      Height          =   1815
      Left            =   840
      TabIndex        =   11
      Top             =   1560
      Visible         =   0   'False
      Width           =   3135
      Begin VB.CommandButton Cerrar 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label pm 
         Caption         =   "0"
         Height          =   255
         Left            =   1080
         TabIndex        =   16
         Top             =   960
         Width           =   855
      End
      Begin VB.Label pf 
         Caption         =   "0"
         Height          =   255
         Left            =   1800
         TabIndex        =   15
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "RECORD:"
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
         TabIndex        =   14
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "TU PUNTUACION:"
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
         TabIndex        =   13
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "FIN DEL JUEGO"
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
         TabIndex        =   12
         Top             =   240
         Width           =   2895
      End
   End
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
      Begin VB.Label pmg 
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
         TabIndex        =   19
         Top             =   240
         Width           =   735
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
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label p 
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
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   4200
      Width           =   4815
      Begin VB.CommandButton Command1 
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
      Begin VB.CommandButton Jugar 
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
   Begin VB.Timer global 
      Interval        =   1
      Left            =   1320
      Top             =   4200
   End
   Begin VB.Timer mov1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   840
      Top             =   4200
   End
   Begin VB.PictureBox figG1 
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
   Begin VB.PictureBox figB1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   600
      Picture         =   "esquivarObjetos.frx":041B
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   1
      Top             =   2880
      Width           =   300
   End
   Begin VB.PictureBox figB1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   3
      Left            =   1680
      Picture         =   "esquivarObjetos.frx":07D1
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   6
      Top             =   2880
      Width           =   300
   End
   Begin VB.PictureBox figB1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   1
      Left            =   2400
      Picture         =   "esquivarObjetos.frx":0B87
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   4
      Top             =   2880
      Width           =   300
   End
   Begin VB.PictureBox figB1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   2
      Left            =   3480
      Picture         =   "esquivarObjetos.frx":0F3D
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   5
      Top             =   2880
      Width           =   300
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
Dim velMov As Integer
Dim X As Integer
Dim pos(4) As Integer
Dim puntaje As Integer
Dim pmax As Integer

Private Sub Cerrar_Click()
info.Visible = False
Jugar.Enabled = True
puntaje = 0
End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub figG1_KeyDown(KeyCode As Integer, Shift As Integer)

    
    If (KeyCode = vbKeyLeft) Then
        figG1.Left = figG1.Left - velMov
    End If
    
    If (KeyCode = vbKeyRight) Then
        figG1.Left = figG1.Left + velMov
    End If



End Sub

Private Sub Form_Load()
pmax = 0
velMov = 200
pos(0) = -2500
pos(1) = -2000
pos(2) = -4500
pos(3) = -7500


figB1(0).Top = pos(0)
figB1(1).Top = pos(1)
figB1(2).Top = pos(2)
figB1(3).Top = pos(3)



End Sub

Private Sub colision()

    If ((figG1.Left < figB1(X).Left + figB1(X).Width) And (figB1(X).Left < figG1.Left + figG1.Width) And (figG1.Top < figB1(X).Top + figB1(X).Height) And (figB1(X).Top < figG1.Top + figG1.Height)) Then
        
        If puntaje >= pmax Then
            
            pmax = puntaje
            pm.Caption = pmax
            MsgBox "Felicitaciones, superaste la puntuación mas alta"
        
        End If
        
        figG1.Enabled = False
        pf.Caption = puntaje
    
        info.Visible = True
        
        figB1(0).Top = pos(0)
        figB1(1).Top = pos(1)
        figB1(2).Top = pos(2)
        figB1(3).Top = pos(3)

        figG1.Left = 2280
        mov1.Enabled = False

    End If
    
End Sub

Private Sub Frame4_DragDrop(Source As Control, X As Single, Y As Single)

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
    
    X = X + 1
    
    If X = 4 Then
    X = 0
    End If
    
    p.Caption = puntaje
    pmg.Caption = pmax
    
    
    

End Sub

Private Sub Jugar_Click()

mov1.Enabled = True
Jugar.Enabled = False
figG1.Enabled = True
figG1.SetFocus

End Sub

Private Sub mov1_Timer()

    figB1(0).Top = figB1(0).Top + 50
    
    If figB1(0).Top >= 4320 Then
    PosfigB1 = (CLng(360 - 4080) * Rnd + 4080)
    figB1(0).Top = 0
    figB1(0).Left = PosfigB1
    puntaje = puntaje + 1
        
    End If
       
       figB1(1).Top = figB1(1).Top + 50
    
    If figB1(1).Top >= 4320 Then
    PosfigB1 = (CLng(360 - 4080) * Rnd + 4080)
    figB1(1).Top = -20
    figB1(1).Left = PosfigB1
     puntaje = puntaje + 1
        
    End If
    
      figB1(2).Top = figB1(2).Top + 50
    
    If figB1(2).Top >= 4320 Then
    PosfigB1 = (CLng(360 - 4080) * Rnd + 4080)
    figB1(2).Top = -20
    figB1(2).Left = PosfigB1
     puntaje = puntaje + 1
        
    End If
    
      figB1(3).Top = figB1(3).Top + 50
    
    If figB1(3).Top >= 4320 Then
    PosfigB1 = (CLng(360 - 4080) * Rnd + 4080)
    figB1(3).Top = -20
    figB1(3).Left = PosfigB1
     puntaje = puntaje + 1
        
    End If
End Sub




