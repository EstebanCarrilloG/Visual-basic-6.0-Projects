VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   Caption         =   "Juego en Visual Basic 6.0"
   ClientHeight    =   7275
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10575
   LinkTopic       =   "Form1"
   ScaleHeight     =   7275
   ScaleWidth      =   10575
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   5640
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
         Left            =   3960
         TabIndex        =   10
         Top             =   120
         Width           =   2055
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "500"
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
            TabIndex        =   11
            Top             =   120
            Width           =   2055
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00000000&
         Caption         =   "Tiempo restante"
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
         TabIndex        =   7
         Top             =   120
         Width           =   1575
         Begin VB.Label L2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "60"
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
            TabIndex        =   8
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
         TabIndex        =   5
         Top             =   120
         Width           =   2055
         Begin VB.Label L1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            TabIndex        =   6
            Top             =   120
            Width           =   1815
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
         Left            =   6120
         TabIndex        =   4
         Top             =   840
         Width           =   4095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Volver a jugar"
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
         Left            =   6120
         TabIndex        =   3
         Top             =   240
         Width           =   4095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Jugar"
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
         Left            =   6120
         TabIndex        =   2
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10335
      Begin VB.Label L3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PRESIONE JUGAR PARA QUE EMPIEZE EL JUEGO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   9
         Top             =   120
         Width           =   10335
      End
      Begin VB.Image D 
         Height          =   735
         Left            =   5400
         Picture         =   "Juego.frx":0000
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   735
      End
      Begin VB.Image C 
         Height          =   1170
         Left            =   4320
         Picture         =   "Juego.frx":07AB
         Stretch         =   -1  'True
         Top             =   2280
         Width           =   1125
      End
   End
   Begin VB.Timer t3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   480
      Top             =   6840
   End
   Begin VB.Timer t1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   120
      Top             =   6840
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UDv As Integer
Dim LRv As Integer
Dim UDr As Integer
Dim LRr As Integer
Dim sum As Integer
Dim TIEMPO As Integer

Private Sub Command1_Click()
TIEMPO = 60
enable

Command1.Visible = False
Command2.Visible = True
L3.Visible = False


End Sub


Private Sub C_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sum = sum + 1
L1.Caption = sum
If sum = 500 Then
MsgBox ("GANASTE")
sum = 0
disable
Command2.Enabled = True



End If
End Sub

Private Sub Command2_Click()
L1.Caption = "0"
L2.Caption = "60"
enable
TIEMPO = 60

Command2.Enabled = False



End Sub

Private Sub D_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
sum = sum - 1



End Sub

Private Sub Form_Load()

C.Enabled = False
D.Enabled = False
Command1.Visible = True
Command2.Visible = False


End Sub

Private Sub Label2_Click()

End Sub

Private Sub salir_Click()
End

End Sub

Private Sub t1_Timer()

UDv = (CLng(120 - 4200) * Rnd + 4200)
LRv = (CLng(0 - 9480) * Rnd + 9480)
D.Top = UDv
D.Left = LRv

UDr = (CLng(120 - 4200) * Rnd + 4200)
LRr = (CLng(0 - 9480) * Rnd + 9480)
C.Top = UDr
C.Left = LRr

End Sub

Private Sub t2_Timer()

End Sub

Private Sub t3_Timer()

TIEMPO = TIEMPO - 1
L2.Caption = TIEMPO

If TIEMPO = 0 Then
disable
Command2.Enabled = True

MsgBox ("PERDISTE,se termino el tiempo")
End If

End Sub

Private Sub encerar()



End Sub

Private Sub enable()
t3.Enabled = True
C.Enabled = True
D.Enabled = True
t1.Enabled = True

End Sub

Private Sub disable()

t3.Enabled = False
t1.Enabled = False
C.Enabled = False
D.Enabled = False
End Sub


