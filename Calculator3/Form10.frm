VERSION 5.00
Begin VB.Form frmLog 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Lod me!"
   ClientHeight    =   2415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7995
   Icon            =   "Form10.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3360
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   4080
      TabIndex        =   4
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load me!"
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000007&
      Caption         =   "Password:"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Calculator 
      BackColor       =   &H80000007&
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
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   6375
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim Response
Dim J As Integer 'loop ctrl
Dim Percentage As Integer



If Not IsNumeric(Not Response) Then Response = 5000
If Response <= 0 Or Response >= 32768 Then Response = 5000

Label2.Visible = True 'Show the blue bit

For J = 1 To Response
    'your event or function being recursed goes here.....
    'or you can put the code below into the recursed function itself.
    'If youre running a fast PC you might want to put something here to slow it down abit, - The progress bar may flicker.

DoEvents
    Percentage = J / Response * 100
    If Percentage <= 0 Then Percentage = 1
    Label2.Width = (Label1.Width / 100 * Percentage) - 30
Calculator.Caption = Str(Percentage) & "% Done"
Next

Label2.Visible = False 'hide the blue bit when done
Calculator.Caption = "Calculator"
Label2.Width = 1
Unload Me











End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Activate()
Label2.Width = 1
Label2.Visible = False
Calculator.Caption = "Calculator"

End Sub

Private Sub Label1_Click()
Unload Me
End Sub

Private Sub Label3_Click()
Unload Me
End Sub

