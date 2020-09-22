VERSION 5.00
Begin VB.Form Shield 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "File Shield"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4800
   Icon            =   "Shield.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Shield.frx":1CCA
   ScaleHeight     =   150
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox TextPassword 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Label LabelStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Initializing... Please Wait"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   4575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Label LabelTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "File Shield"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "Shield"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public WithEvents Encoder As EncrypticEncoder
Attribute Encoder.VB_VarHelpID = -1

Sub InitiateFileShield()
    Dim FileNum As Long, ED As Byte, EM As Byte, EY As Integer, ExpirationDate As Date
    Me.Show
    
    FileNum = FreeFile
    Open App.Path & "\" & App.EXEName & ".exe" For Binary Access Read As #FileNum
        ReDim ThisFile(LOF(FileNum) - 1)
        Get #FileNum, 1, ThisFile()
        If LOF(FileNum) = SEFSize Then
            LabelStatus.Caption = "File data not available"
            Close #FileNum
            Exit Sub
        End If
    Close #FileNum
    
    Details = GetFileData(ThisFile)
    
    App.Title = GetString(Details.Title)
    Shield.Caption = GetString(Details.Title)
    LabelTitle.Caption = GetString(Details.Title)
    
    On Error GoTo DateError::
    CopyMemory ByVal VarPtr(ED), ByVal VarPtr(Details.ExpirationDate), 1
    CopyMemory ByVal VarPtr(EM), ByVal VarPtr(Details.ExpirationDate) + 1, 1
    CopyMemory ByVal VarPtr(EY), ByVal VarPtr(Details.ExpirationDate) + 2, 2
    If Not (ED = 0 And EM = 0 And EY = 0) Then
        ExpirationDate = CDate(CStr(ED) & " " & MonthName(CLng(EM)) & " " & CStr(EY))
        If Now >= ExpirationDate Then GoTo DateError::
    End If
    
    LabelStatus.Caption = "Authorization Required"
    TextPassword.Visible = True
    Label2.Visible = True
    
    Exit Sub
DateError::
    LabelStatus.Caption = "Data Has Expired"
End Sub

Private Sub Encoder_ReturnProgress(ProgPercent As Long, ProgDone As Long, ProgTotal As Long)
    LabelStatus.Caption = "Processing: " & CStr(ProgPercent) & "%"
End Sub

Private Sub TextPassword_KeyPress(KeyAscii As Integer)
    Dim Pass() As Byte
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Len(TextPassword.Text) > 0 Then
            GeneratePasswordData TextPassword.Text, Pass
            If GetUniqueID(Pass) = GetString(Details.PasswordUID) Then
                TextPassword.Enabled = False
                If WriteTemporaryFile(TextPassword.Text) Then
                    LabelStatus.Caption = "Extraction Complete"
                    Me.Hide
                    DoEvents
                    OpenApplication OpenFileName, GetString(Details.CommandLine)
                    DoEvents
                    Kill OpenFileName
                    Unload Me
                Else
                    LabelStatus.Caption = "Extraction Error"
                End If
            Else
                LabelStatus.Caption = "Access Denied"
                TextPassword.Enabled = False
                Timer1.Enabled = True
            End If
        Else
            LabelStatus.Caption = "Password required"
        End If
    ElseIf KeyAscii = 27 Then
        KeyAscii = 0
        TextPassword.Text = ""
    End If
End Sub

Private Sub Timer1_Timer()
    TextPassword.Text = ""
    TextPassword.Enabled = True
    Timer1.Enabled = False
    LabelStatus.Caption = "Authorization Required"
End Sub
