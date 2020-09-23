VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000007&
   Caption         =   "Form2"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8565
   LinkTopic       =   "Form2"
   ScaleHeight     =   5595
   ScaleWidth      =   8565
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picScroll 
      BackColor       =   &H80000012&
      Height          =   3015
      Left            =   120
      Picture         =   "frmW.frx":0000
      ScaleHeight     =   2955
      ScaleWidth      =   8235
      TabIndex        =   0
      Top             =   840
      Width           =   8295
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "Evripides Kyriakou"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   855
      Left            =   1920
      TabIndex        =   1
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub
