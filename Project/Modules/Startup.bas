Attribute VB_Name = "Startup"
Option Explicit

Sub Main()
    Set Shield.Encoder = New EncrypticEncoder
    
    Shield.Show
    DoEvents
    Shield.InitiateFileShield
End Sub
