VERSION 5.00
Begin VB.Form frmDigital 
   BackColor       =   &H80000006&
   Caption         =   "Digital Calculator"
   ClientHeight    =   4455
   ClientLeft      =   165
   ClientTop       =   255
   ClientWidth     =   5085
   Icon            =   "frmDigital.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   5085
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command23 
      Caption         =   "Cos"
      Height          =   375
      Left            =   3960
      TabIndex        =   30
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Tan"
      Height          =   375
      Left            =   3960
      TabIndex        =   29
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Sin"
      Height          =   375
      Left            =   3960
      TabIndex        =   28
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Deg"
      Height          =   375
      Left            =   3960
      TabIndex        =   27
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command19 
      Caption         =   "+/-"
      Height          =   375
      Left            =   3960
      TabIndex        =   26
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command18 
      Caption         =   "%"
      Height          =   375
      Left            =   3960
      TabIndex        =   25
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command17 
      Caption         =   "."
      Height          =   495
      Left            =   2760
      TabIndex        =   21
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton Command16 
      Caption         =   "C"
      Height          =   375
      Left            =   600
      TabIndex        =   17
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton Command15 
      Caption         =   "0"
      Height          =   495
      Left            =   1320
      TabIndex        =   15
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton Command14 
      Caption         =   "="
      Height          =   495
      Left            =   600
      TabIndex        =   14
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton Command13 
      Caption         =   "/"
      Height          =   495
      Left            =   3120
      TabIndex        =   13
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton Command12 
      Caption         =   "*"
      Height          =   495
      Left            =   2280
      TabIndex        =   12
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton Command11 
      Caption         =   "-"
      Height          =   495
      Left            =   3120
      TabIndex        =   11
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton Command10 
      Caption         =   "+"
      Height          =   495
      Left            =   2280
      TabIndex        =   10
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton Command9 
      Caption         =   "9"
      Height          =   495
      Left            =   600
      TabIndex        =   9
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      Caption         =   "8"
      Height          =   495
      Left            =   3120
      TabIndex        =   8
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "7"
      Height          =   495
      Left            =   2280
      TabIndex        =   7
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "6"
      Height          =   495
      Left            =   1440
      TabIndex        =   6
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "5"
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "4"
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "3"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      Height          =   495
      Left            =   1440
      TabIndex        =   2
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000007&
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   4575
   End
   Begin VB.Label Label17 
      BackColor       =   &H80000008&
      Caption         =   "0"
      Height          =   375
      Left            =   3720
      TabIndex        =   23
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      Caption         =   "0"
      Height          =   375
      Left            =   3840
      TabIndex        =   22
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000008&
      Caption         =   "0"
      Height          =   375
      Left            =   3960
      TabIndex        =   20
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000007&
      Caption         =   "0"
      Height          =   495
      Left            =   3960
      TabIndex        =   19
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000007&
      Caption         =   "0"
      Height          =   375
      Left            =   4080
      TabIndex        =   18
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000008&
      Height          =   375
      Left            =   3000
      TabIndex        =   16
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Menu q 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu a 
         Caption         =   "Scientific"
      End
      Begin VB.Menu t 
         Caption         =   "Exit"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu p 
      Caption         =   "Help"
      Visible         =   0   'False
      Begin VB.Menu o 
         Caption         =   "Help me!"
      End
      Begin VB.Menu h 
         Caption         =   "About Calculator"
      End
   End
End
Attribute VB_Name = "frmDigital"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub a_Click()
Command18.Visible = True
Command19.Visible = True
Command20.Visible = True
Command21.Visible = True
Command22.Visible = True
Command23.Visible = True
End Sub

Private Sub Command1_Click()
Label17 = 0
Text1.SelText = CDec(1)

End Sub

Private Sub Command10_Click()
On Error Resume Next
Label17 = 0
If Text1 = Label3 Then
    Text1 = ""
    Label1 = 0
ElseIf Label1 = 0 Then
    Label1 = 1
    Label2 = 0 + Label2 + Text1
    Text1 = ""
    Label1 = 0

End If
End Sub

Private Sub Command11_Click()
On Error Resume Next
Label17 = 0
If Text1 = Label3 Then
    Label1 = 1
    Text1 = ""
ElseIf Label1 = 0 Then
    Label2 = 0 + Label2 + Text1
    Label1 = 1
    Text1 = ""
    Label1 = 1
ElseIf Label1 = 2 Then
    Label2 = Label2 * Text1
    Label1 = 1
    Text1 = ""
ElseIf Label1 = 1 Then
    Label2 = Label2 - Text1
    Text1 = ""
ElseIf Label1 = 3 Then
    Label2 = Label2 / Text1
    Label1 = 1
    Text1 = ""
End If
End Sub

Private Sub Command12_Click()
On Error Resume Next
Label17 = 0
If Text1 = Label3 Then
    Label1 = 2
    Text1 = ""
ElseIf Label1 = 0 Then
    Label2 = 0 + Label2 + Text1
    Label1 = 2
    Text1 = ""
ElseIf Label1 = 1 Then
    Label2 = Label2 - Text1
    Label1 = 2
    Text1 = ""
ElseIf Label1 = 2 Then
    Label2 = Label2 * Text1
    Text1 = ""
ElseIf Label1 = 3 Then
    Label2 = Label2 / Text1
    Label1 = 2
    Text1 = ""
End If
End Sub

Private Sub Command13_Click()
On Error Resume Next
Label17 = 0
If Text1 = Label3 Then
    Label1 = 3
    Text1 = ""
ElseIf Label1 = 0 Then
    Label2 = 0 + Label2 + Text1
    Label1 = 3
    Text1 = ""
ElseIf Label1 = 1 Then
    Label2 = Label2 - Text1
    Label1 = 3
    Text1 = ""
ElseIf Label1 = 2 Then
    Label2 = Label2 * Text1
    Label1 = 3
    Text1 = ""
ElseIf Label1 = 3 Then
    Label2 = Label2 / Text1
    Text1 = ""
End If
End Sub

Private Sub Command14_Click()
On Error Resume Next
Dim i As Integer
If Label17 = 1 Then
    Label2 = Label2 + 0
ElseIf Label1 = 0 Then
    Label2 = 0 + Label2 + Text1
    Label3 = Label2
    Label1 = 0
    Label17 = 1
    Text1 = Label3
ElseIf Label1 = 1 Then
    Label2 = Label2 - Text1
    Label3 = Label2
    Label1 = 0
    Text1 = Label3
    Label17 = 1
ElseIf Label1 = 2 Then
    Label2 = Label2 * Text1
    Label3 = Label2
    Label1 = 0
    Text1 = Label3
    Label17 = 1
ElseIf Label1 = 3 Then
    Label2 = CDec(Label2) / CDec(Text1)
    Label3 = Label2
    Label1 = 0
    Text1 = Label3
    Label17 = 1
End If
Label8 = 1
End Sub

Private Sub Command15_Click()
Label17 = 0
Text1.SelText = CDec(0)

End Sub

Private Sub Command16_Click()
On Error Resume Next
If Label12 = 1 Then
    Label1 = 0
    Label3 = 0
    Label2 = 0
    Label17 = 0
    Label10 = 0
    Text1 = ""
ElseIf Label12 = 0 Then
    Label1 = 0
    Label2 = 0
    Label3 = 0
    Label17 = 0
    Label10 = 0
    Text1 = ""
End If
End Sub

Private Sub Command17_Click()
Label17 = 0
Text1.SelText = "."
End Sub

Private Sub Command18_Click()
On Error Resume Next
' This performs a percentage calculation
Label17 = 0
Text1 = Text1 / 100
Text1 = Text1 * 100
Text1.Text = Text1.Text + "%"
End Sub

Private Sub Command19_Click()
' Use this to toggle plus and minus
On Error Resume Next
Label17 = 0
If Label9.Caption = 0 Then
    Text1 = Text1 - Text1 - Text1
    Label9 = 1
ElseIf Label9 = 1 Then
    Text1 = Text1 - Text1 - Text1
    Label9 = 0
End If
End Sub

Private Sub Command2_Click()
Label17 = 0
Text1.SelText = CDec(2)

End Sub

Private Sub Command20_Click()
' A procedure to convert a number to decimal
On Error Resume Next
Label17 = 0
Text1 = Text1 / 100
End Sub

Private Sub Command21_Click()
' This works out the sin of a number
On Error Resume Next
Label17 = 0
Text1 = Text1 * Pi / 180
Text1 = Sin(Text1)
End Sub

Private Sub Command22_Click()
' This works out the tangent of a number
On Error Resume Next
Label17 = 0
Text1 = Text1 * Pi / 180
Text1 = Tan(Text1)
End Sub

Private Sub Command23_Click()
' This works out the cosine of a number
On Error Resume Next
Label17 = 0
Text1 = Text1 * Pi / 180
Text1 = Cos(Text1)
End Sub

Private Sub Command3_Click()
Label17 = 0
Text1.SelText = CDec(3)

End Sub

Private Sub Command4_Click()
Label17 = 0
Text1.SelText = CDec(4)

End Sub

Private Sub Command5_Click()
Label17 = 0
Text1.SelText = CDec(5)

End Sub

Private Sub Command6_Click()
Label17 = 0
Text1.SelText = CDec(6)

End Sub

Private Sub Command7_Click()
Label17 = 0
Text1.SelText = CDec(7)

End Sub

Private Sub Command8_Click()
Label17 = 0
Text1.SelText = CDec(8)

End Sub

Private Sub Command9_Click()
Label17 = 0
Text1.SelText = CDec(9)

End Sub

Private Sub Form_Load()
Command18.Visible = False
Command19.Visible = False
Command20.Visible = False
Command21.Visible = False
Command22.Visible = False
Command23.Visible = False
End Sub

Private Sub h_Click()
frmHelp4.Show 1
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
q.Visible = True
    p.Visible = True
    frmDigital.Height = 5340
End Sub

Private Sub o_Click()
frmHelp.Show 1
End Sub

Private Sub t_Click()
Unload Me
End Sub

