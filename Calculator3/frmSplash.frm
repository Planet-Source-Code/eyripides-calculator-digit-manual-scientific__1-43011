VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H80000007&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4305
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7440
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   7440
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Timer Timer1 
         Interval        =   3000
         Left            =   3480
         Top             =   2040
      End
      Begin VB.Image imgLogo 
         Height          =   1785
         Left            =   360
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   795
         Width           =   1695
      End
      Begin VB.Label lblCopyright 
         BackColor       =   &H80000012&
         Caption         =   "Copyright:2002-2003"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   3120
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         BackColor       =   &H80000012&
         Caption         =   "Production E.K"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   3360
         Width           =   2415
      End
      Begin VB.Label lblWarning 
         BackColor       =   &H80000012&
         Caption         =   "Warning:All rights reserved"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   3660
         Width           =   6855
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "V.1.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   5880
         TabIndex        =   4
         Top             =   3000
         Width           =   555
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Calculator"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   900
         Left            =   2520
         TabIndex        =   6
         Top             =   1140
         Width           =   3105
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackColor       =   &H80000008&
         Caption         =   "E.K"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   525
         Left            =   3240
         TabIndex        =   5
         Top             =   720
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub Frame1_Click()
Unload Me
End Sub

Private Sub imgLogo_Click()
Unload Me
End Sub

Private Sub lblCompany_Click()
Unload Me
End Sub

Private Sub lblCompanyProduct_Click()
Unload Me
End Sub

Private Sub lblCopyright_Click()
Unload Me
End Sub

Private Sub lblProductName_Click()
Unload Me
End Sub

Private Sub lblVersion_Click()
Unload Me
End Sub

Private Sub lblWarning_Click()
Unload Me
End Sub

Private Sub Picture1_GotFocus()

End Sub

Private Sub Timer1_Timer()
Unload Me
End Sub

