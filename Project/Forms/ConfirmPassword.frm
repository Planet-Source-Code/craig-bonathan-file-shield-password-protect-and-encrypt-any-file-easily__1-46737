VERSION 5.00
Begin VB.Form ConfirmPassword 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Confirm Password"
   ClientHeight    =   855
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4935
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   4935
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Confirm"
      Default         =   -1  'True
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox TextPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      MaxLength       =   250
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "ConfirmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Password As String
Public State As Long

Private Sub Command1_Click()
    Password = TextPassword.Text
    State = 2
End Sub

Private Sub Command2_Click()
    Password = ""
    State = 1
End Sub
