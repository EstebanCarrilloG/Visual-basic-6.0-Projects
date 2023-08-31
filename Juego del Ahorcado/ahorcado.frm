VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "JUEGO DEL AHORCADO"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   5775
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
      Width           =   5295
      Begin VB.CommandButton vj 
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
         Width           =   5055
      End
      Begin VB.Label infoNdigit 
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   23
         Top             =   480
         Width           =   5055
      End
      Begin VB.Label Estado 
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
         Width           =   5055
      End
   End
   Begin VB.Frame ilp 
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
      Width           =   5295
      Begin VB.TextBox palabra 
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
         Width           =   3735
      End
      Begin VB.CommandButton start 
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
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command1 
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
      Width           =   5295
   End
   Begin VB.Timer juego 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5880
      Top             =   6480
   End
   Begin VB.Frame pbra 
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
      Width           =   5295
      Begin VB.TextBox J 
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
      Begin VB.TextBox J 
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
      Begin VB.TextBox J 
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
      Begin VB.TextBox J 
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
      Begin VB.TextBox J 
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
      Begin VB.TextBox J 
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
      Begin VB.TextBox J 
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
      Begin VB.TextBox J 
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
      Begin VB.TextBox J 
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
      Begin VB.TextBox J 
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
      Begin VB.TextBox J 
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
      Begin VB.TextBox J 
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
      Begin VB.TextBox J 
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
      Begin VB.TextBox J 
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
      Width           =   5295
      Begin VB.Line N11 
         BorderColor     =   &H0080C0FF&
         BorderStyle     =   3  'Dot
         BorderWidth     =   5
         Visible         =   0   'False
         X1              =   3960
         X2              =   4440
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Shape N5 
         BorderColor     =   &H000000FF&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   4080
         Shape           =   2  'Oval
         Top             =   1560
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Shape N4 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   4320
         Shape           =   3  'Circle
         Top             =   1320
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Shape N3 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   3960
         Shape           =   3  'Circle
         Top             =   1320
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Line N10 
         BorderWidth     =   10
         Visible         =   0   'False
         X1              =   4320
         X2              =   4680
         Y1              =   2520
         Y2              =   3480
      End
      Begin VB.Line N9 
         BorderWidth     =   10
         Visible         =   0   'False
         X1              =   4080
         X2              =   3720
         Y1              =   2400
         Y2              =   3480
      End
      Begin VB.Line N8 
         BorderWidth     =   10
         Visible         =   0   'False
         X1              =   4560
         X2              =   4920
         Y1              =   1920
         Y2              =   2520
      End
      Begin VB.Line N7 
         BorderWidth     =   10
         Visible         =   0   'False
         X1              =   3840
         X2              =   3360
         Y1              =   1920
         Y2              =   2520
      End
      Begin VB.Shape N6 
         FillStyle       =   0  'Solid
         Height          =   975
         Left            =   3840
         Shape           =   4  'Rounded Rectangle
         Top             =   1800
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Shape N2 
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   3720
         Shape           =   3  'Circle
         Top             =   960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Line N1 
         BorderColor     =   &H0080C0FF&
         BorderStyle     =   3  'Dot
         BorderWidth     =   5
         Visible         =   0   'False
         X1              =   4200
         X2              =   4200
         Y1              =   600
         Y2              =   1800
      End
      Begin VB.Line Line1 
         BorderWidth     =   15
         X1              =   2280
         X2              =   1320
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
   End
   Begin VB.Shape Shape5 
      Height          =   6855
      Left            =   120
      Top             =   240
      Width           =   5535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CANT As Integer
Dim I As Integer
Dim B As Integer
Dim XX As Integer
Dim PA As String
Dim LE As String
Dim LE2 As String
Dim P As Integer
Dim G As Integer
Dim N As Integer

Private Sub Command1_Click()
End
End Sub


Function Verificar_Tecla(Tecla_Presionada)
    
    
Dim Teclas As String
    
    
    Teclas = "ABCDEFGHIJKLMNÑOPQRSTUVWXYZ" & Chr(vbKeyBack)
    
    If InStr(1, Teclas, Chr(Tecla_Presionada)) Then
        
        Verificar_Tecla = Tecla_Presionada
    Else
        
        Verificar_Tecla = 0
    End If
    

End Function
Private Sub Form_Load()

    Estado.Caption = "Ingrese una palabra y presione Jugar"

End Sub

Private Sub juego_Timer()


  I = 0
  N = 0

  Do While I < CANT + 100
    
    LE = InputBox("INGRESE UNA LETRA, RECUERDE QUE SOLO SE ADMITEN LETRAS MAYUSCULAS.", "Ingreso", "", 0, 0)
    XX = 0
    
      For B = 1 To CANT
      
        LE2 = Mid(PA, B, 1)
        
         If LE = LE2 Then
         
            Estado.Caption = "Letra correcta: " & LE
            
               Select Case B
               
                 Case 1
                   J(0).Text = LE
                 Case 2
                   J(1).Text = LE
                 Case 3
                   J(2).Text = LE
                 Case 4
                   J(3).Text = LE
                 Case 5
                   J(4).Text = LE
                 Case 6
                   J(5).Text = LE
                 Case 7
                   J(6).Text = LE
                 Case 8
                   J(7).Text = LE
                 Case 9
                   J(8).Text = LE
                 Case 10
                   J(9).Text = LE
                 Case 11
                   J(10).Text = LE
                 Case 12
                   J(11).Text = LE
                 Case 13
                   J(12).Text = LE
                 Case 14
                   J(13).Text = LE

               End Select
               
            XX = 1
            
         End If
      Next B


      If XX = 0 Then
      
        N = N + 1
        Estado.Caption = "Letra incorrecta: " & LE
        
      End If
        
      Select Case N
      
       Case 1
         N1.Visible = True
         N11.Visible = True
       Case 2
         N2.Visible = True
       Case 3
         N6.Visible = True
       Case 4
         N7.Visible = True
       Case 5
         N8.Visible = True
       Case 6
         N9.Visible = True
       Case 7
         N10.Visible = True
    
      End Select
    
      G = 0
      
      For P = 0 To CANT - 1
      
        If J(P) <> "" Then
    
          G = G + 1
        End If
      Next P
    
      If G = CANT Then
      
        juego.Enabled = False
        MsgBox "Ganaste"
        Estado.Caption = ""
        PA = ""
        I = CANT + 100
        vj.Visible = True
        
      End If
    
    
    
      If N = 8 Then
        juego.Enabled = False
        MsgBox "Perdiste"
        Estado.Caption = "PALABRA INGRESADA: " & PA
        PA = ""
        I = CANT + 100
        vj.Visible = True
      End If
    
    
      I = I + 1
  
  Loop

End Sub

Private Sub palabra_KeyPress(KeyAscii As Integer)
KeyAscii = Verificar_Tecla(KeyAscii)

End Sub

Private Sub start_Click()

 For I = 0 To 13
  J(I).Text = ""
  Next I

 PA = palabra.Text
 CANT = Len(PA)
    
 
 If (PA = "") Then
 
  Estado.Caption = "Ingrese una palabra."
  
 Else
 
  If (CANT < 4) Then
    Estado.Caption = "Palabra muy corta"
    ElseIf CANT > 13 Then
    Estado.Caption = "Superaste el numero de letras permitidas"
    CANT = 0
    
  Else
  
    infoNdigit.Caption = "La palabra ingresada tiene " & CANT & " letras"
    ilp.Visible = False
    pbra.Visible = True
    Estado.Caption = "Comienza el juego"
    juego.Enabled = True
    
    For X = 0 To CANT - 1
    J(X).Visible = True
    
    Next
                           
  End If
 End If


End Sub


Private Sub vj_Click()

    palabra.Text = ""
    ilp.Visible = True
    vj.Visible = False
    Estado.Caption = "Ingrese una palabra y presione Jugar"
    infoNdigit.Caption = ""
    N = 0

    N1.Visible = False
    N2.Visible = False
    N3.Visible = False
    N4.Visible = False
    N5.Visible = False
    N6.Visible = False
    N7.Visible = False
    N8.Visible = False
    N9.Visible = False
    N10.Visible = False
    N11.Visible = False

End Sub
