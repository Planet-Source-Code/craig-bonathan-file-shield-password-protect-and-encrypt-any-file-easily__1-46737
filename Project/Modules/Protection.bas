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

Public Type ProtectionDetails_Type
    ExpirationDate As Long
    CommandLine As String * 251
    Extension As String * 11
    Title As String * 51
    PasswordUID As String * 48
    ApplicationUID As String * 48
End Type

Function EndString(Text As String) As String
    EndString = Text & Chr(0)
End Function

Function SetFileData(Data() As Byte, Details As ProtectionDetails_Type, Password As String)
    Dim Length As Long, Temp() As Byte, Pass() As Byte
    Length = UBound(Data) + 1
    ReDim Preserve Data(Length + SEFSize + LenB(Details) - 1)
    
    GeneratePasswordData Password, Pass
    
    ' Copy and encode file data, offsetting the data by SEFSize
    ReDim Temp(Length)
    CopyMemory ByVal VarPtr(Temp(0)), ByVal VarPtr(Data(0)), Length
    ProtectForm.Encoder.EncodeData Temp, Pass, True
    CopyMemory ByVal VarPtr(Data(SEFSize)), ByVal VarPtr(Temp(0)), Length
    
    ' Copy SEF data to beginning of file data
    CopyMemory ByVal VarPtr(Data(0)), ByVal VarPtr(SEFData(0)), SEFSize
    
    ' Copy and encode protection details to the end of the file
    ReDim Temp(LenB(Details) - 1)
    CopyMemory ByVal VarPtr(Temp(0)), ByVal VarPtr(Details), LenB(Details)
    ProtectForm.Encoder.EncodeData Temp, SEFData, True
    CopyMemory ByVal VarPtr(Data(UBound(Data) - LenB(Details) + 1)), ByVal VarPtr(Temp(0)), LenB(Details)
End Function

Function GeneratePasswordData(Password As String, Data() As Byte)
    Data = StrConv(Password, vbFromUnicode)
End Function
