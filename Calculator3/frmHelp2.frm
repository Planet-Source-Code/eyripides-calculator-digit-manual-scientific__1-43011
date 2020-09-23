VERSION 5.00
Begin VB.Form frmHelp2 
   BackColor       =   &H8000000D&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ÂïÞèåéá!"
   ClientHeight    =   30
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   945
   Icon            =   "frmHelp2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   30
   ScaleWidth      =   945
   ShowInTaskbar   =   0   'False
   Begin VB.Timer T2 
      Left            =   6000
      Top             =   600
   End
   Begin VB.Timer T1 
      Left            =   5040
      Top             =   600
   End
   Begin VB.CommandButton Command5 
      Caption         =   "/"
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "*"
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "+"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "C"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000006&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "¸îïäïò"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   3000
      MouseIcon       =   "frmHelp2.frx":0442
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000D&
      Caption         =   "ÂïÞèåéá!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4080
      TabIndex        =   10
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000D&
      Caption         =   "Äéáéñåóç"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   9
      Top             =   3240
      Width           =   3375
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000D&
      Caption         =   "Ðïëëáðëáóéáóìïò"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   8
      Top             =   2520
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      Caption         =   "Áöáßñåóç"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   7
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "Ðñüóèåóç"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   6
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "Êáèáñéóìüò ÕðïëïãéóôéêÞò"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   5
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "frmHelp2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

'As you can see I haven't changed the following values
'from the properties so that every one can understand
'what this project needs to work. If you want you can
'change them directly from the properties.
T1.Interval = 5
'You can change the values
'directly from properties
T2.Enabled = False
T2.Interval = 5
End Sub

Private Sub Label7_Click()
T2.Enabled = True
End Sub

Private Sub Label8_Click()

End Sub


Private Sub T1_Timer()
'The following commands are responsible for all the work.

'The first 2(two) lines make the form grow bigger,
Me.Height = Me.Height + 30
Me.Width = Me.Width + 30

'And the these 2(two) lines make the form move
Me.Left = Me.Left + 20
Me.Top = Me.Top + 20

'This command disables T1(Timer1) when Me.Height goes to 6000
If Me.Height = 6000 Then T1.Enabled = False
'Here you can put width too, it doesn't matter we only
'want a value where our T1 commands will stop being repeated
End Sub

Private Sub T2_Timer()
'Me.Height goes - 200
Me.Height = Me.Height - 200

'Until Me.Height reaches 510
'where it unloads
If Me.Height = 510 Then Unload Me
End Sub
