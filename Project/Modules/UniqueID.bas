Attribute VB_Name = "UniqueID"
' Copyright Craig Bonathan 2003

Option Explicit

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
     Destination As Any, _
     Source As Any, _
     ByVal Length As Long)

Function GetUniqueID(Data() As Byte) As String
    Dim DataSize As Long, Mean As Single, SD As Single
    Dim Totals1() As Long, Totals2() As Long, ByteOrder1 As Single, ByteOrder2 As Single
    Dim TempMean As Single
    Dim ByteFrequency(255) As Long, TotalByteValue As Long
    Dim Pos As Long, Total As Long, Pos2 As Long
    Dim UniqueIDData() As Byte
    Dim OriginalSize As Long
    OriginalSize = UBound(Data)
    
    DataSize = OriginalSize + 1
    
    For Pos = 0 To DataSize - 1
        ByteFrequency(Data(Pos)) = ByteFrequency(Data(Pos)) + 1
        TotalByteValue = TotalByteValue + Data(Pos)
        TotalByteValue = TotalByteValue Mod 256 ^ 3
        ByteFrequency(Data(Pos)) = ByteFrequency(Data(Pos)) Mod 256
    Next
    Mean = TotalByteValue / DataSize
    
    For Pos = 0 To 255
        Total = Total + ByteFrequency(Pos) * (Pos - Mean) ^ 2
        Total = Total Mod (256 ^ 3)
    Next
    SD = Total / DataSize
    SD = Sqr(SD)
    
    If (DataSize Mod 10) > 0 Then DataSize = DataSize - (DataSize Mod 10) + 10
    ReDim Preserve Data(DataSize - 1)
    ReDim Totals1(5)
    ReDim Totals2(9)
    
    For Pos2 = 0 To 4
        For Pos = Pos2 To DataSize - 1 Step 5
            Totals1(Pos2) = Totals1(Pos2) + Data(Pos)
            Totals1(Pos2) = Totals1(Pos2) Mod 256 ^ 3
        Next
    Next
    Total = 0
    
    For Pos = 0 To UBound(Totals1)
        Total = Total + Totals1(Pos)
        Totals1(Pos) = Totals1(Pos) Mod 256
    Next
    TempMean = Total / 5
    TempMean = TempMean Mod 256
    Total = 0
    
    For Pos = 0 To 4
        Total = Total + Totals1(Pos) * (Pos - TempMean) ^ 2
        Total = Total Mod 256 ^ 3
    Next
    ByteOrder1 = Total / DataSize
    ByteOrder1 = Sqr(ByteOrder1)
    
    For Pos2 = 0 To 9
        For Pos = Pos2 To DataSize - 1 Step 10
            Totals2(Pos2) = Totals2(Pos2) + Data(Pos)
        Next
    Next
    Total = 0
    
    For Pos = 0 To UBound(Totals2)
        Total = Total + Totals2(Pos)
        Totals2(Pos) = Totals2(Pos) Mod 256
    Next
    TempMean = Total / 10
    TempMean = TempMean Mod 256
    Total = 0
    
    For Pos = 0 To 9
        Total = Total + Totals2(Pos) * (Pos - TempMean) ^ 2
        Total = Total Mod 256 ^ 3
    Next
    ByteOrder2 = Total / DataSize
    ByteOrder2 = Sqr(ByteOrder2)
    
    ReDim UniqueIDData(19)
    CopyMemory ByVal VarPtr(UniqueIDData(0)), ByVal VarPtr(DataSize), 4
    CopyMemory ByVal VarPtr(UniqueIDData(4)), ByVal VarPtr(Mean), 4
    CopyMemory ByVal VarPtr(UniqueIDData(8)), ByVal VarPtr(SD), 4
    CopyMemory ByVal VarPtr(UniqueIDData(12)), ByVal VarPtr(ByteOrder1), 4
    CopyMemory ByVal VarPtr(UniqueIDData(16)), ByVal VarPtr(ByteOrder2), 4
    
    For Pos = 0 To 19
        UniqueIDData(Pos) = (UniqueIDData(Pos) + (Pos * (UniqueIDData(Pos) Xor Pos))) Mod 256
    Next
    GetUniqueID = GetUniqueHex(UniqueIDData())
    GetUniqueID = SplitUniqueID(GetUniqueID, 5)
    
    ReDim Preserve Data(OriginalSize)
End Function

Private Function GetUniqueHex(Data() As Byte) As String
    Dim Pos As Long, Temp As String
    For Pos = 0 To UBound(Data)
        Temp = Hex(Data(Pos))
        If Len(Temp) = 1 Then Temp = (Hex(Pos Mod 16)) & Temp
        GetUniqueHex = GetUniqueHex & Temp
    Next
End Function

Private Function SplitUniqueID(Data As String, Interval As Long) As String
    Dim Pos As Long
    For Pos = 1 To Len(Data) Step Interval
        SplitUniqueID = SplitUniqueID & Mid(Data, Pos, Interval) & "-"
    Next
    SplitUniqueID = Mid(SplitUniqueID, 1, Len(SplitUniqueID) - 1)
End Function
