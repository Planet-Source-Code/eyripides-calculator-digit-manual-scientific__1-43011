VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000007&
   Caption         =   "Calculator"
   ClientHeight    =   7785
   ClientLeft      =   2475
   ClientTop       =   1860
   ClientWidth     =   9855
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   9855
   Begin VB.Timer T3 
      Left            =   840
      Top             =   1680
   End
   Begin VB.Timer Timer1 
      Left            =   5040
      Top             =   1680
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Why Me?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7080
      TabIndex        =   14
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Help!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7080
      MouseIcon       =   "Form1.frx":0442
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Off"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4560
      TabIndex        =   12
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   11
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton cmdMultiply 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   10
      Top             =   6480
      Width           =   735
   End
   Begin VB.CommandButton cmdDivision 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   9
      Top             =   6480
      Width           =   855
   End
   Begin VB.CommandButton cmdSubtract 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   8
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton cmdAddition 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   7
      Top             =   5760
      Width           =   855
   End
   Begin VB.TextBox txtSecond 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000003&
      Height          =   615
      Left            =   4320
      TabIndex        =   5
      Top             =   3480
      Width           =   3015
   End
   Begin VB.TextBox txtFirst 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000003&
      Height          =   615
      Left            =   4320
      TabIndex        =   4
      Top             =   2400
      Width           =   3015
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000007&
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   9855
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   7800
      TabIndex        =   16
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000007&
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lblMessage 
      BackColor       =   &H80000012&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   21.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   4320
      TabIndex        =   6
      Top             =   4440
      Width           =   3015
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000008&
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000007&
      Caption         =   "Write the second number"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      Caption         =   "Write the first number"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000008&
      Caption         =   "   Calculator"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1095
      Left            =   1800
      TabIndex        =   0
      Top             =   600
      Width           =   5895
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu a 
         Caption         =   "Digital Calculator"
      End
      Begin VB.Menu i 
         Caption         =   "Scientific"
      End
      Begin VB.Menu r 
         Caption         =   "Exit"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Visible         =   0   'False
      Begin VB.Menu w 
         Caption         =   "Help me!"
      End
      Begin VB.Menu l 
         Caption         =   "About Calculator..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Private Sub a_Click()
frmDigital.Show 1
End Sub

Private Sub cmdAddition_Click()
Dim cFirst As Currency
Dim cSecond As Currency
Dim cResult As Currency
cFirst = Val(txtFirst.Text)
cSecond = Val(txtSecond.Text)
cResult = cFirst + cSecond
lblMessage.Caption = cResult
End Sub

Private Sub cmdClear_Click()
txtFirst = ""
txtSecond = ""
lblMessage = ""
txtFirst.SetFocus
End Sub

Private Sub cmdExit_Click()
T3.Enabled = True
End Sub

Private Sub cmdMultiply_Click()
Dim cFirst As Currency
Dim cSecond As Currency
Dim cResult As Currency
cFirst = Val(txtFirst.Text)
cSecond = Val(txtSecond.Text)
cResult = cFirst * cSecond
lblMessage.Caption = cResult
End Sub

Private Sub cmdDivision_Click()
Dim cFirst As Currency
Dim cSecond As Currency
Dim cResult As Currency
cFirst = Val(txtFirst.Text)
cSecond = Val(txtSecond.Text)
cResult = cFirst / cSecond
lblMessage.Caption = cResult
End Sub

Private Sub cmdSubtract_Click()
Dim cFirst As Currency
Dim cSecond As Currency
Dim cResult As Currency
cFirst = Val(txtFirst.Text)
cSecond = Val(txtSecond.Text)
cResult = cFirst - cSecond
lblMessage.Caption = cResult
End Sub

Private Sub Command1_Click()
frmHelp.Show 1
Timer1.Interval = 1
End Sub

Private Sub Command2_Click()
frmHelp4.Show 1
End Sub

Private Sub f_Click(Index As Integer)

End Sub




Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Form_Load()
frmSplash.Show 1
frmLog.Show 1
Timer1.Interval = 1
'You can change the values
'directly from properties
T3.Enabled = False
T3.Interval = 5
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
Me.PopupMenu Me.mnuFile
Me.PopupMenu Me.mnuHelp
End If
End Sub

Private Sub i_Click()
MsgBox "It will be added in a new version!", vbInformation, "Info..."
End Sub

Private Sub l_Click()
frmHelp4.Show 1
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 mnuFile.Visible = True
    mnuHelp.Visible = True
    frmMain.Height = 8295
End Sub

Private Sub m_Click(Index As Integer)
Unload Me
End Sub

Private Sub mnu_Click(Index As Integer)
frmHelp.Show 1
End Sub

Private Sub mnu1_Click(Index As Integer)
frmHelp4.Show 1
End Sub

Private Sub r_Click()
Unload Me
End Sub



Private Sub T3_Timer()
'Me.Height goes - 200
Me.Height = Me.Height - 200

'Until Me.Height reaches 510
'where it unloads
If Me.Height = 510 Then Unload Me
End Sub

Private Sub Timer1_Timer()
Label5.Caption = Time ' this shows the time
Label6.Caption = Date ' and this shows the date
End Sub



Private Sub w_Click()
frmHelp.Show 1
End Sub
