VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   3735
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   4755
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox objeto 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   2040
      Picture         =   "mover objeto.frx":0000
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   0
      Top             =   1560
      Width           =   450
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
      Top             =   3240
      Width           =   4815
   End
   Begin VB.Shape Shape2 
      FillStyle       =   7  'Diagonal Cross
      Height          =   4455
      Left            =   4440
      Top             =   -720
      Width           =   375
   End
   Begin VB.Shape Shape1 
      FillStyle       =   7  'Diagonal Cross
      Height          =   4455
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

Dim desplazamiento As Integer

Private Sub objeto_KeyDown(KeyCode As Integer, Shift As Integer)

If (KeyCode = vbKeyUp) Then
    objeto.Top = objeto.Top - desplazamiento
ElseIf (KeyCode = vbKeyDown) Then
    objeto.Top = objeto.Top + desplazamiento
ElseIf (KeyCode = vbKeyLeft) Then
    objeto.Left = objeto.Left - desplazamiento
ElseIf (KeyCode = vbKeyRight) Then
    objeto.Left = objeto.Left + desplazamiento
End If

'limites del formulario

If (objeto.Left <= 360) Then
    objeto.Left = 360
End If

If (objeto.Left >= 3960) Then
    objeto.Left = 3960
End If

If (objeto.Top <= 480) Then
    objeto.Top = 480
End If

If (objeto.Top >= 2760) Then
    objeto.Top = 2760
End If



End Sub

Private Sub Form_Load()
    desplazamiento = 500 'Variable utilizada para controlar la velocidad de desplazamiento variando la distancia(rango) de movimiento
End Sub

