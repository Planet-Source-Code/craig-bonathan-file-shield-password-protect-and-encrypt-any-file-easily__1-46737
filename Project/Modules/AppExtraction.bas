Attribute VB_Name = "AppExtraction"
Option Explicit

Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function WriteFile Lib "kernel32.dll" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Overlapped) As Long
Private Declare Function ShellExecuteEx Lib "shell32.dll" (lpExecInfo As SHELLEXECUTEINFO) As Long
Private Declare Function WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Public OpenFileName As String

Private Type Overlapped
    ternal As Long
    ternalHigh As Long
    offset As Long
    OffsetHigh As Long
    hEvent As Long
End Type

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    ' fields
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Function OpenApplication(FileName As String, CommandLine As String)
    Dim ShellExecInfo As SHELLEXECUTEINFO
    
    ShellExecInfo.lpFile = FileName
    ShellExecInfo.lpDirectory = App.Path
    ShellExecInfo.lpParameters = CommandLine
    ShellExecInfo.nShow = 1
    ShellExecInfo.fMask = &H40
    ShellExecInfo.lpVerb = "Open"
    ShellExecInfo.hwnd = Shield.hwnd
    ShellExecInfo.cbSize = Len(ShellExecInfo)
    
    ShellExecuteEx ShellExecInfo
    
    WaitForSingleObject ShellExecInfo.hProcess, &HFFFFFFFF
    
    CloseHandle ShellExecInfo.hProcess
End Function

Function WriteTemporaryFile(Password As String) As Boolean
    Dim NotUsed1 As SECURITY_ATTRIBUTES, NotUsed2 As Overlapped
    Dim FileName As String, FileNumber As Long, FileHandle As Long
    Dim FSO As Object
    Dim Pass() As Byte, UID As String
    GeneratePasswordData Password, Pass
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    For FileNumber = 0 To 65535
        FileName = App.Path & "\Temp" & Hex(FileNumber) & GetString(Details.Extension)
        If Not FSO.FileExists(FileName) Then
            Exit For
        Else
            FileName = ""
        End If
    Next
    
    If FileName = "" Then Exit Function
    
    Shield.Encoder.DecodeData ThisFile, Pass, True
    
    If GetString(Details.ApplicationUID) <> "" Then
        Shield.LabelStatus.Caption = "Checking File UID"
        DoEvents
        UID = GetUniqueID(ThisFile)
        If UID <> GetString(Details.ApplicationUID) Then Exit Function
    End If
            
    FileHandle = CreateFile(FileName & Chr(0), &H40000000, 0, NotUsed1, 1, 0, 0)
    WriteFile FileHandle, ThisFile(0), UBound(ThisFile) + 1, 0, NotUsed2
    CloseHandle FileHandle
    
    OpenFileName = FileName
    WriteTemporaryFile = True
End Function
