VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SEMAFOROS "
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   8490
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   720
      Top             =   3720
   End
   Begin VB.Shape luzVerdeDerecha 
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   6000
      Top             =   2280
      Width           =   735
   End
   Begin VB.Shape luzAmarillaDerecha 
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   6000
      Top             =   1440
      Width           =   735
   End
   Begin VB.Shape luzRojaDerecha 
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   6000
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape luzRojaIzquierda 
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   1800
      Top             =   600
      Width           =   735
   End
   Begin VB.Shape luzAmarillaIzquierda 
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   1800
      Top             =   1440
      Width           =   735
   End
   Begin VB.Shape luzVerdeIzquierda 
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   1800
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   3195
      Left            =   120
      Picture         =   "semaforo.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   4140
   End
   Begin VB.Image Image2 
      Height          =   3195
      Left            =   4320
      Picture         =   "semaforo.frx":64AE
      Stretch         =   -1  'True
      Top             =   120
      Width           =   4140
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer

Private Sub Form_Load()
End Sub

Private Sub Timer1_Timer()
    If a = 0 Then
      luzVerdeIzquierda.Visible = False
      luzRojaDerecha.Visible = False
      End If
    If a = 4 Then
      luzVerdeIzquierda.Visible = True
      luzRojaDerecha.Visible = False
      luzAmarillaIzquierda.Visible = False
     End If
    If a = 5 Then
      luzRojaIzquierda.Visible = False
      luzAmarillaIzquierda.Visible = True
      luzVerdeDerecha.Visible = False
      luzRojaDerecha.Visible = True
    End If
    If a = 7 Then
      luzRojaIzquierda.Visible = False
      luzAmarillaDerecha.Visible = False
      luzVerdeDerecha.Visible = True
    End If
    If a = 9 Then
       luzRojaDerecha.Visible = False
       luzAmarillaDerecha.Visible = True
       luzVerdeIzquierda.Visible = False
       luzRojaIzquierda.Visible = True
    End If
    
    a = a + 1
    
    If a = 10 Then
        a = 0
    End If
    
End Sub
