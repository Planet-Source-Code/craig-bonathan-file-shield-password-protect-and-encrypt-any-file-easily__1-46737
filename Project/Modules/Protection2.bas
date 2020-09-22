Attribute VB_Name = "Protection"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub MoveMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)


' IMPORTANT! SEFSize MUST be set to the EXACT same size as Lock.exe (in bytes), otherwise neither program will work.
' There is a constant in the protection module in the other project which must be set to the same value.
' To find this value, compile Lock.exe, check its size, and then enter it here. Then compile both projects.
' When you modify anything in FileShield_Lock, it may change the compiled file size, so check it.
Public Const SEFSize As Long = 200704

Public SEFData(SEFSize - 1) As Byte
Public Const SEFUID As String = ""
Public Const SEFName As String = "\Lock.exe"

Public ThisFile() As Byte
Public Details As ProtectionDetails_Type

Public Type ProtectionDetails_Type
    ExpirationDate As Long
    CommandLine As String * 251
    Extension As String * 11
    Title As String * 51
    PasswordUID As String * 48
    ApplicationUID As String * 48
End Type

Function GetString(Text As String) As String
    Dim Pos As Long
    Pos = InStr(1, Text, Chr(0))
    If Pos = 0 Then GetString = Text
    If Pos = 1 Then GetString = ""
    If Pos > 1 Then GetString = Mid(Text, 1, Pos - 1)
End Function

Function GetFileData(Data() As Byte) As ProtectionDetails_Type
    Dim Temp() As Byte
    ' Copy SEF data
    CopyMemory ByVal VarPtr(SEFData(0)), ByVal VarPtr(Data(0)), SEFSize
    
    ' Copy and decode file details
    ReDim Temp(LenB(GetFileData))
    CopyMemory ByVal VarPtr(Temp(0)), ByVal VarPtr(Data(UBound(Data) - LenB(GetFileData) + 1)), _
            LenB(GetFileData)
    Shield.Encoder.DecodeData Temp, SEFData, True
    CopyMemory ByVal VarPtr(GetFileData), ByVal VarPtr(Temp(0)), LenB(GetFileData)
    
    ' Move file data to beggining and resize
    MoveMemory ByVal VarPtr(Data(0)), ByVal VarPtr(Data(SEFSize)), UBound(Data) - SEFSize - LenB(GetFileData) + 1
    ReDim Preserve Data(UBound(Data) - SEFSize - LenB(GetFileData))
End Function

Function GeneratePasswordData(Password As String, Data() As Byte)
    Data = StrConv(Password, vbFromUnicode)
End Function

Function DecryptDataFile(Data() As Byte, Password As String)
    Dim Pass() As Byte
    GeneratePasswordData Password, Pass
    Shield.Encoder.DecodeData Data, Pass, True
End Function
