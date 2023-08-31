VERSION 5.00
Begin VB.Form Juego 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Juego en Visual Basic 6.0"
   ClientHeight    =   7815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer t2 
      Interval        =   50
      Left            =   960
      Top             =   7800
   End
   Begin VB.Frame Frame6 
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
      Begin VB.OptionButton n2 
         BackColor       =   &H00404040&
         Caption         =   "Nivel 2"
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
         Height          =   255
         Left            =   4800
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton n3 
         BackColor       =   &H00000000&
         Caption         =   "Nivel 3"
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
         Height          =   255
         Left            =   6240
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton n1 
         BackColor       =   &H00404040&
         Caption         =   "Nivel 1"
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
         Height          =   255
         Left            =   3480
         TabIndex        =   12
         Top             =   240
         Width           =   1215
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
         Begin VB.Label Objetivo 
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
         Begin VB.Label L2 
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
         Begin VB.Label L1 
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
      Begin VB.CommandButton jugar 
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10335
      Begin VB.Image E 
         Height          =   1095
         Left            =   1800
         Picture         =   "Juego 05 01 2021.frx":0000
         Stretch         =   -1  'True
         Top             =   3360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label L3 
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
      Begin VB.Image D 
         Height          =   735
         Left            =   5400
         Picture         =   "Juego 05 01 2021.frx":07AB
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   735
      End
      Begin VB.Image C 
         Height          =   1170
         Left            =   4320
         Picture         =   "Juego 05 01 2021.frx":0F56
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

Dim sum As Integer
Dim TIEMPO As Integer
Dim n As Integer

Private Sub C_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sum = sum + 1
L1.Caption = sum

If sum = Objetivo Then
sum = 0
MsgBox "Felicitaciones, Completaste el nivel.", vbInformation, "Ganaste"

disable


n1.Enabled = True
n2.Enabled = True
n3.Enabled = True
jugar.Enabled = True
t2.Enabled = True
t3.Enabled = False



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

L1.Caption = sum

End Sub

Private Sub E_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
sum = sum - 1

L1.Caption = sum
End Sub

Private Sub Form_Load()

C.Enabled = False
D.Enabled = False
jugar.Visible = True



End Sub

Private Sub Label2_Click()

End Sub

Private Sub Option2_Click()

End Sub

Private Sub jugar_Click()

t2.Enabled = False
TIEMPO = L2

enable
jugar.Enabled = False
L3.Visible = False

n1.Enabled = False
n2.Enabled = False
n3.Enabled = False

L1.Caption = "0"


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

UDr1 = (CLng(120 - 4200) * RnU + 4200)
LRr1 = (CLng(0 - 9480) * Rnd + 9480)
E.Top = UDr1
E.Left = LRr1



End Sub

Private Sub t2_Timer()

If n1.Value = True Then
jugar.Enabled = True
L2 = 60
L1 = 0
Objetivo = 500
t1.Interval = 700
E.Visible = False

End If

If n2.Value = True Then
jugar.Enabled = True
L2 = 50
L1 = 0
Objetivo = 700
t1.Interval = 600
E.Visible = False
End If



If n3.Value = True Then
jugar.Enabled = True
L2 = 40
L1 = 0
Objetivo = 1000
t1.Interval = 500
E.Visible = True

End If


End Sub

Private Sub t3_Timer()

TIEMPO = TIEMPO - 1
L2.Caption = TIEMPO

If TIEMPO = 0 Then
disable
sum = 0
jugar.Enabled = True
MsgBox "Se termino el tiempo", 64, "Perdiste"

n1.Enabled = True
n2.Enabled = True
n3.Enabled = True
t2.Enabled = True
jugar.Enabled = True

End If

End Sub

Private Sub encerar()



End Sub

Private Sub enable()
t3.Enabled = True
C.Enabled = True
D.Enabled = True
E.Enabled = True
t1.Enabled = True

End Sub

Private Sub disable()

t3.Enabled = False
t1.Enabled = False
C.Enabled = False
D.Enabled = False
E.Enabled = False
End Sub


