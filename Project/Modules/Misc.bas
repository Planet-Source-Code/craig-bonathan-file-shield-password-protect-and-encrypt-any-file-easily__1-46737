Attribute VB_Name = "Misc"
Option Explicit

Public LogFile As Long

Function OpenLog()
    LogFile = FreeFile
    Open App.Path & "\log.txt" For Output As #LogFile
End Function

Function CloseLog()
    Close #LogFile
End Function

Function ShowLog(Text As String)
    LogForm.Show
    DoEvents
    LogForm.TextLog.Text = LogForm.TextLog.Text & Text & vbCrLf & vbCrLf
    Print #LogFile, CStr(Now) & ": " & Text
End Function

Function HideLog()
    LogForm.Hide
End Function

Function LoadSecondaryExecutable() As Boolean
    Dim FileNum As Long
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If Not FSO.FileExists(App.Path & SEFName) Then Exit Function
    FileNum = FreeFile
    Open App.Path & SEFName For Binary Access Read As #FileNum
        Get #FileNum, 1, SEFData()
    Close #FileNum
    LoadSecondaryExecutable = True
End Function
