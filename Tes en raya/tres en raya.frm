VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Tres en raya"
   ClientHeight    =   5715
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleMode       =   0  'User
   ScaleWidth      =   8520
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton b1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4680
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   360
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
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
      TabIndex        =   19
      Top             =   1200
      Width           =   4335
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   4815
      Left            =   4560
      TabIndex        =   9
      Top             =   120
      Width           =   3855
      Begin VB.CommandButton b5 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   1320
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1440
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton b6 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   2520
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1440
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton b2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   1320
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton b3 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   2520
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton b4 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1440
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton b7 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2640
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton b8 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   1320
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2640
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton b9 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   2520
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2640
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton INICIAR 
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
   Begin VB.CommandButton Command1 
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
      Top             =   4080
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
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   120
      TabIndex        =   5
      Top             =   1680
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
   Begin VB.CommandButton restart 
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
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   6840
      Top             =   6480
   End
   Begin VB.TextBox j2 
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
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox j1 
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
      Top             =   240
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
      Top             =   840
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
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label mar 
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
      Top             =   5040
      Width           =   8295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sw As Integer
Dim name1 As String
Dim name2 As String
Dim X As Integer
Dim marp1 As Integer
Dim marp2 As Integer

Private Sub limpiar()
mensaje


b1.Enabled = False
b2.Enabled = False
b3.Enabled = False
b4.Enabled = False
b5.Enabled = False
b6.Enabled = False
b7.Enabled = False
b8.Enabled = False
b9.Enabled = False

INICIAR.Enabled = True
INICIAR.Caption = "VOLVER A JUGAR"

End Sub

Private Sub b1_Click()
If sw = 0 Then
b1.Caption = "O"
b1.Enabled = False
sw = 1
comprovar
Else
b1.Caption = "X"
b1.Enabled = False
sw = 0
comprovar
End If









End Sub

Private Sub b2_Click()
If sw = 0 Then
b2.Caption = "O"
b2.Enabled = False
sw = 1
comprovar
Else
b2.Caption = "X"
b2.Enabled = False
sw = 0
comprovar
End If


End Sub

Private Sub b3_Click()
If sw = 0 Then
b3.Caption = "O"
b3.Enabled = False
sw = 1
comprovar
Else
b3.Caption = "X"
b3.Enabled = False
sw = 0
comprovar
End If


End Sub

Private Sub b4_Click()
If sw = 0 Then
b4.Caption = "O"
b4.Enabled = False
sw = 1
comprovar
Else
b4.Caption = "X"
b4.Enabled = False
sw = 0
comprovar
End If



End Sub

Private Sub b5_Click()
If sw = 0 Then
b5.Caption = "O"
b5.Enabled = False
sw = 1
comprovar
Else
b5.Caption = "X"
b5.Enabled = False
sw = 0
comprovar

End If




End Sub

Private Sub b6_Click()
If sw = 0 Then
b6.Caption = "O"
b6.Enabled = False
sw = 1
comprovar
Else
b6.Caption = "X"
b6.Enabled = False
sw = 0
comprovar
End If



End Sub

Private Sub b7_Click()
If sw = 0 Then
b7.Caption = "O"
b7.Enabled = False
sw = 1
comprovar
Else
b7.Caption = "X"
b7.Enabled = False
sw = 0
comprovar
End If


End Sub

Private Sub b8_Click()
If sw = 0 Then
b8.Caption = "O"
b8.Enabled = False
sw = 1
comprovar
Else
b8.Caption = "X"
b8.Enabled = False
sw = 0
comprovar
End If



End Sub

Private Sub b9_Click()
If sw = 0 Then
b9.Caption = "O"
b9.Enabled = False
sw = 1
comprovar
Else
b9.Caption = "X"
b9.Enabled = False
sw = 0
comprovar
End If





End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub Command6_Click()

End Sub

Private Sub Command5_Click()

End Sub

Private Sub Command2_Click()
name1 = j1.Text
name2 = j2.Text




If name1 = "" And name2 = "" Then
mar.Caption = "Ingrese los nombres de los jugadores"

End If

If name1 <> "" And name2 <> "" Then
mar.Caption = "Presione jugar para continuar "
INICIAR.Enabled = True
Timer1.Enabled = False
j1.Enabled = False
j2.Enabled = False
End If
If name1 = "" Then
 mar.Caption = "Ingrese el nombre del Jugador 1"
End If

If name2 = "" Then
 mar.Caption = "Ingrese el nombre del Jugador 2"
End If




b1.Enabled = False
b2.Enabled = False
b3.Enabled = False
b4.Enabled = False
b5.Enabled = False
b6.Enabled = False
b7.Enabled = False
b8.Enabled = False
b9.Enabled = False
End Sub

Private Sub Form_Load()

b1.Enabled = False
b2.Enabled = False
b3.Enabled = False
b4.Enabled = False
b5.Enabled = False
b6.Enabled = False
b7.Enabled = False
b8.Enabled = False
b9.Enabled = False



Randomize
  sw = (CLng(0 - 1) * Rnd + 1)

End Sub

Private Sub comprovar()

If b1.Caption = "O" And b2.Caption = "O" And b3.Caption = "O" Or b4.Caption = "O" And b5.Caption = "O" And b6.Caption = "O" Or b7.Caption = "O" And b8.Caption = "O" And b9.Caption = "O" Or b1.Caption = "O" And b4.Caption = "O" And b7.Caption = "O" Or b2.Caption = "O" And b5.Caption = "O" And b8.Caption = "O" Or b3.Caption = "O" And b6.Caption = "O" And b9.Caption = "O" Or b1.Caption = "O" And b5.Caption = "O" And b9.Caption = "O" Or b3.Caption = "O" And b5.Caption = "O" And b7.Caption = "O" Then
marp1 = marp1 + 1
mar.Caption = name1 & " ES EL GANADOR"
limpiar
ElseIf b1.Caption = "X" And b2.Caption = "X" And b3.Caption = "X" Or b4.Caption = "X" And b5.Caption = "X" And b6.Caption = "X" Or b7.Caption = "X" And b8.Caption = "X" And b9.Caption = "X" Or b1.Caption = "X" And b4.Caption = "X" And b7.Caption = "X" Or b2.Caption = "X" And b5.Caption = "X" And b8.Caption = "X" Or b3.Caption = "X" And b6.Caption = "X" And b9.Caption = "X" Or b1.Caption = "X" And b5.Caption = "X" And b9.Caption = "X" Or b3.Caption = "X" And b5.Caption = "X" And b7.Caption = "X" Then
marp2 = marp2 + 1
mar.Caption = name2 & " ES EL GANADOR"
limpiar
Else
If b1.Caption <> "" And b2.Caption <> "" And b3.Caption <> "" And b4.Caption <> "" And b5.Caption <> "" And b6.Caption <> "" And b7.Caption <> "" And b8.Caption <> "" And b1.Caption <> "" And b9.Caption <> "" Then
mar.Caption = "ES UN EMPATE"
Randomize
    sw = (CLng(0 - 1) * Rnd + 1)

limpiar
End If
End If


End Sub


Private Sub INICIAR_Click()
b1.Caption = ""
b2.Caption = ""
b3.Caption = ""
b4.Caption = ""
b5.Caption = ""
b6.Caption = ""
b7.Caption = ""
b8.Caption = ""
b9.Caption = ""


j1.Enabled = False
j2.Enabled = False




mar.Caption = ""
     
mensaje
If sw = 0 Then
MsgBox name1 & " Empieza", vbInformation, "TRES EN RAYA"
Else
MsgBox name2 & " Empieza", vbInformation, "TRES EN RAYA"
End If


b1.Enabled = True
b2.Enabled = True
b3.Enabled = True
b4.Enabled = True
b5.Enabled = True
b6.Enabled = True
b7.Enabled = True
b8.Enabled = True
b9.Enabled = True
INICIAR.Enabled = False


End Sub

Private Sub mensaje()
info.Caption = name1 & ": " & marp1 & vbCrLf & name2 & ": " & marp2
End Sub

Private Sub Label3_Click()

End Sub

Private Sub restart_Click()

b1.Enabled = False
b2.Enabled = False
b3.Enabled = False
b4.Enabled = False
b5.Enabled = False
b6.Enabled = False
b7.Enabled = False
b8.Enabled = False
b9.Enabled = False

Timer1.Enabled = True
j1.Text = ""
j2.Text = ""
j1.Enabled = True
j2.Enabled = True

marp1 = 0
marp2 = 0

info.Caption = ""
mar.Caption = "Ingrese los nombres de los jugadores"
INICIAR.Caption = "JUGAR"



End Sub



Private Sub j1_KeyPress(KeyAscii As Integer)
    KeyAscii = Verificar_Tecla(KeyAscii)
End Sub
Private Sub j2_KeyPress(KeyAscii As Integer)
    KeyAscii = Verificar_Tecla(KeyAscii)
End Sub

Function Verificar_Tecla(Tecla_Presionada)
    
    
Dim Teclas As String
    
    
    Teclas = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZabcdefghijklmnñopqrstuvwxyz" & Chr(vbKeyBack)
    
    If InStr(1, Teclas, Chr(Tecla_Presionada)) Then
        
        Verificar_Tecla = Tecla_Presionada
    Else
        
        Verificar_Tecla = 0
    End If
    

End Function

Private Sub Timer1_Timer()
INICIAR.Enabled = False
End Sub
