VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form ProtectForm 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Shield"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5175
   Icon            =   "ProtectForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   454
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   345
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog FileDialog 
      Left            =   4560
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Select a File to Protect"
   End
   Begin VB.CommandButton ButtonExecute 
      Caption         =   "Execute"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   6000
      Width           =   4935
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Security"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   120
      TabIndex        =   6
      Top             =   4440
      Width           =   4935
      Begin VB.TextBox TextExpires 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3360
         TabIndex        =   11
         Top             =   360
         Width           =   1455
      End
      Begin VB.CheckBox CheckExpires 
         BackColor       =   &H00000000&
         Caption         =   "Expires:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2400
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
      Begin VB.CheckBox CheckGFUID 
         BackColor       =   &H00000000&
         Caption         =   "Generate File UID"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox TextPassword 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         MaxLength       =   250
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   960
         Width           =   4695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Password:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "General"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   4935
      Begin VB.TextBox TextCommandLine 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         MaxLength       =   250
         TabIndex        =   13
         Top             =   1800
         Width           =   4695
      End
      Begin VB.CommandButton ButtonBrowse 
         Caption         =   "Browse"
         Height          =   375
         Left            =   3840
         TabIndex        =   5
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox TextFileName 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   3615
      End
      Begin VB.TextBox TextTitle 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   2
         Top             =   600
         Width           =   4695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "Command Line:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "File:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Title:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "File Shield"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   1320
      Width           =   4935
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Copyright Craig Bonathan 2003"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1680
      Width           =   4935
   End
   Begin VB.Image Image1 
      Height          =   1050
      Left            =   360
      Picture         =   "ProtectForm.frx":1CCA
      Top             =   120
      Width           =   4500
   End
   Begin VB.Label LabelStatus 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Not Busy"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   6480
      Width           =   4935
   End
End
Attribute VB_Name = "ProtectForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public WithEvents Encoder As EncrypticEncoder
Attribute Encoder.VB_VarHelpID = -1
Dim FileData() As Byte

Private Sub ButtonBrowse_Click()
    On Error GoTo Skip::
    FileDialog.DialogTitle = "Select a file to protect"
    FileDialog.FileName = ""
    FileDialog.Flags = &H4 Or &H1000 Or &H200000
    FileDialog.Filter = ""
    FileDialog.ShowOpen
    TextFileName.Text = FileDialog.FileName
Skip::
End Sub

Private Sub ButtonExecute_Click()
    Dim FileNum As Long, Details As ProtectionDetails_Type, Pass() As Byte
    Dim Temp As String, Pos As Long, SaveAs As String, ExitRoutine As Boolean
    Dim ExpirationDate As Date, ED As Byte, EM As Byte, EY As Integer
    Dim FSO As Object
    
    If TextFileName.Text = "" Then
        LabelStatus.Caption = "Please select a file"
        Exit Sub
    End If
    If TextTitle.Text = "" Then
        LabelStatus.Caption = "Please enter a title"
        Exit Sub
    End If
    If TextPassword.Text = "" Then
        LabelStatus.Caption = "Please enter a password"
        Exit Sub
    End If
    
    If CheckExpires.Value = 1 Then
        On Error GoTo DateError::
        ExpirationDate = CDate(TextExpires.Text)
        On Error GoTo 0
        ED = Day(ExpirationDate)
        EM = Month(ExpirationDate)
        EY = Year(ExpirationDate)
        CopyMemory ByVal VarPtr(Details.ExpirationDate), ByVal VarPtr(ED), 1
        CopyMemory ByVal VarPtr(Details.ExpirationDate) + 1, ByVal VarPtr(EM), 1
        CopyMemory ByVal VarPtr(Details.ExpirationDate) + 2, ByVal VarPtr(EY), 2
    End If
    
    On Error GoTo Skip::
    FileDialog.DialogTitle = "Select a destination for the executable"
    FileDialog.FileName = ""
    FileDialog.Flags = &H4 Or &H2 Or &H200000
    FileDialog.Filter = "Application (*.exe)|*.exe"
    FileDialog.ShowSave
    SaveAs = FileDialog.FileName
    On Error GoTo 0
    
    ProtectForm.Enabled = False
    ConfirmPassword.State = 0
    ConfirmPassword.Password = ""
    ConfirmPassword.Show
    Do Until ConfirmPassword.State <> 0
        DoEvents
    Loop
    If ConfirmPassword.State = 1 Then
        LabelStatus = "Process Aborted"
        ExitRoutine = True
    Else
        If ConfirmPassword.Password = "" Then
            LabelStatus = "Process Aborted"
            ExitRoutine = True
        ElseIf ConfirmPassword.Password <> TextPassword.Text Then
            LabelStatus = "Password Failure"
            ExitRoutine = True
        End If
    End If
    Unload ConfirmPassword
    ProtectForm.Enabled = True
    ProtectForm.Show
    If ExitRoutine = True Then Exit Sub
    
    Frame1.Enabled = False
    Frame2.Enabled = False
    ButtonExecute.Enabled = False
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If Not FSO.FileExists(TextFileName.Text) Then
        LabelStatus.Caption = "File does not exist"
        Exit Sub
    End If
    
    ShowLog "Loading """ & TextFileName.Text & """ in to memory"
    LabelStatus.Caption = "Loading File"
    
    FileNum = FreeFile
    Open TextFileName.Text For Binary Access Read As #FileNum
        If LOF(FileNum) = 0 Then
            Close #FileNum
            ShowLog "Error. File is empty. Process Aborted."
            LabelStatus.Caption = "File is empty"
            Frame1.Enabled = True
            Frame2.Enabled = True
            ButtonExecute.Enabled = True
            Exit Sub
        End If
        ReDim FileData(LOF(FileNum) - 1)
        Get #FileNum, 1, FileData()
    Close #FileNum
    
    If CheckGFUID.Value = 1 Then
        ShowLog "Generating Application Unique ID"
        LabelStatus.Caption = "Generating Application UID"
        DoEvents
        Details.ApplicationUID = EndString(GetUniqueID(FileData))
    End If
    
    ShowLog "Formatting Data Arrays"
    
    GeneratePasswordData TextPassword.Text, Pass
    Details.PasswordUID = EndString(GetUniqueID(Pass))
    Details.Title = EndString(TextTitle.Text)
    Details.CommandLine = EndString(TextCommandLine.Text)
    
    Temp = TextFileName.Text
    Pos = InStrRev(Temp, ".")
    If Pos = 0 Then
        Temp = ""
    Else
        Temp = Mid(Temp, Pos)
    End If
    Details.Extension = EndString(Temp)
    
    SetFileData FileData(), Details, TextPassword.Text
    
    ShowLog "Writing Executable to """ & SaveAs & """"
    
    LabelStatus.Caption = "Writing File"
    
    FileNum = FreeFile
    Open SaveAs For Binary Access Write As #FileNum
    Close #FileNum
    
    Kill SaveAs
    
    FileNum = FreeFile
    Open SaveAs For Binary Access Write As #FileNum
        Put #FileNum, 1, FileData()
    Close #FileNum
    
    Frame1.Enabled = True
    Frame2.Enabled = True
    ButtonExecute.Enabled = True
    LabelStatus.Caption = "Process Complete"
    ShowLog "Executable Written"
    HideLog
Skip::
    Exit Sub
DateError::
    LabelStatus.Caption = "Invalid Date"
End Sub

Private Sub Encoder_ReturnProgress(ProgPercent As Long, ProgDone As Long, ProgTotal As Long)
    LabelStatus.Caption = "Processing: " & CStr(ProgPercent) & "%"
End Sub

Private Sub Form_Load()
    Set Encoder = New EncrypticEncoder
    Me.Show
    DoEvents
    OpenLog
    ShowLog "Log initialized."
    ShowLog "Loading secondary executable..."
    If LoadSecondaryExecutable = True Then
        ShowLog "Loading complete."
    Else
        ShowLog "Error. Secondary executable (""" & App.Path & SEFName & """) not found."
        MsgBox ("Critical Error. See log.txt for details.")
        End
    End If
    HideLog
    Frame1.Enabled = True
    Frame2.Enabled = True
    ButtonExecute.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ShowLog "Closing log..."
    CloseLog
    End
End Sub
