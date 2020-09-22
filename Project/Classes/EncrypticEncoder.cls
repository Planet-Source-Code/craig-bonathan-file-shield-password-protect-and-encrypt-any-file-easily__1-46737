VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EncrypticEncoder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Encryptic Encoder
' Copyright Craig Bonathan 2003

Option Explicit

' Events
Public Event ReturnProgress(ProgPercent As Long, ProgDone As Long, ProgTotal As Long)

' Single byte encoding / decoding functions
Private Function EncryptByte1(DataByte As Byte, KeyByte As Byte) As Byte
    Dim Temp As Long
    Temp = CLng(DataByte)
    Temp = Temp + KeyByte
    If Temp > 255 Then Temp = Temp - 256
    EncryptByte1 = CByte(Temp)
    EncryptByte1 = EncryptByte1 Xor KeyByte
End Function
Private Function DecryptByte1(DataByte As Byte, KeyByte As Byte) As Byte
    Dim Temp As Long
    DataByte = DataByte Xor KeyByte
    Temp = CLng(DataByte)
    Temp = Temp - KeyByte
    If Temp < 0 Then Temp = Temp + 256
    DecryptByte1 = CByte(Temp)
End Function
Private Function ByteStage2(DataByte As Byte, ByteNo As Long) As Byte
    Dim Temp As Long
    Temp = ByteNo Mod 256
    Temp = 255 - (255 Xor Temp)
    ByteStage2 = CByte(Temp) Xor DataByte
End Function


' Data encoding / decoding functions
Public Function EncodeData(ByRef DataArray() As Byte, ByRef KeyArray() As Byte, _
        DoProgress As Boolean, Optional ProgDoneOvr As Long, Optional ProgTotalOvr As Long)
    Dim DataLength As Long
    Dim KeyLength As Long
    Dim KeyPos As Long
    Dim DataPos As Long
    DataLength = UBound(DataArray)
    KeyLength = UBound(KeyArray)
    If DataLength = 0 Or KeyLength = 0 Then Exit Function
    For DataPos = 0 To DataLength
        KeyPos = ((DataPos / KeyLength) - Int(DataPos / KeyLength)) * KeyLength
        KeyPos = KeyLength - KeyPos
        DataArray(DataPos) = ByteStage2(DataArray(DataPos), DataPos)
        DataArray(DataPos) = EncryptByte1(DataArray(DataPos), KeyArray(KeyPos))
        If DoProgress = True Then
            If DataLength < 100 Then
                If ProgTotalOvr > 0 Then
                    RaiseEvent ReturnProgress(100, (ProgDoneOvr + DataPos), ProgTotalOvr)
                Else
                    RaiseEvent ReturnProgress(100, DataPos, DataLength)
                End If
            Else
                If DataPos Mod (DataLength / 100) = 0 Then
                    If ProgTotalOvr > 0 Then
                        RaiseEvent ReturnProgress(((ProgDoneOvr + DataPos) / ProgTotalOvr) * 100, _
                                (ProgDoneOvr + DataPos), ProgTotalOvr)
                    Else
                        RaiseEvent ReturnProgress((DataPos / DataLength) * 100, DataPos, DataLength)
                    End If
                End If
            End If
        End If
        DoEvents
    Next
End Function
Public Function DecodeData(ByRef DataArray() As Byte, ByRef KeyArray() As Byte, _
        DoProgress As Boolean, Optional ProgDoneOvr As Long, Optional ProgTotalOvr As Long)
    Dim DataLength As Long
    Dim KeyLength As Long
    Dim KeyPos As Long
    Dim DataPos As Long
    DataLength = UBound(DataArray)
    KeyLength = UBound(KeyArray)
    If DataLength = 0 Or KeyLength = 0 Then Exit Function
    For DataPos = 0 To DataLength
        KeyPos = ((DataPos / KeyLength) - Int(DataPos / KeyLength)) * KeyLength
        KeyPos = KeyLength - KeyPos
        DataArray(DataPos) = DecryptByte1(DataArray(DataPos), KeyArray(KeyPos))
        DataArray(DataPos) = ByteStage2(DataArray(DataPos), DataPos)
        If DoProgress = True Then
            If DataLength < 100 Then
                If ProgTotalOvr > 0 Then
                    RaiseEvent ReturnProgress(100, (ProgDoneOvr + DataPos), ProgTotalOvr)
                Else
                    RaiseEvent ReturnProgress(100, DataPos, DataLength)
                End If
            Else
                If DataPos Mod (DataLength / 100) = 0 Then
                    If ProgTotalOvr > 0 Then
                        RaiseEvent ReturnProgress(((ProgDoneOvr + DataPos) / ProgTotalOvr) * 100, _
                                (ProgDoneOvr + DataPos), ProgTotalOvr)
                    Else
                        RaiseEvent ReturnProgress((DataPos / DataLength) * 100, DataPos, DataLength)
                    End If
                End If
            End If
        End If
        DoEvents
    Next
    
    Exit Function
End Function


' Class events
Private Sub Class_Initialize()
    ' Do nothing
End Sub

Private Sub Class_Terminate()
    ' Do nothing
End Sub
