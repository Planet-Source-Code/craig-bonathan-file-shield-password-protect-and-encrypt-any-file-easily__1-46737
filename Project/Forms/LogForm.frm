VERSION 5.00
Begin VB.Form LogForm 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Execution Log"
   ClientHeight    =   2895
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   3975
   Icon            =   "LogForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   3975
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TextLog 
      Height          =   2655
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "LogForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    TextLog.Left = 0
    TextLog.Top = 0
    TextLog.Width = Me.ScaleWidth
    TextLog.Height = Me.ScaleHeight
End Sub

Private Sub Form_Resize()
    TextLog.Left = 0
    TextLog.Top = 0
    TextLog.Width = Me.ScaleWidth
    TextLog.Height = Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Hide
    Cancel = True
End Sub
