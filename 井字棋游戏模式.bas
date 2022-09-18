Attribute VB_Name = "Module1"
Option Explicit

Enum GameModeType
    SoloFirst = 1
    SoloSecond = 2
    Duo = 3
    Self = 4
End Enum

Public GameMode As GameModeType
Public IsCancelled As Boolean

Sub SwitchOrder()
    If GameMode = SoloFirst Then
        GameMode = SoloSecond
    ElseIf GameMode = SoloSecond Then
        GameMode = SoloFirst
    End If
End Sub
