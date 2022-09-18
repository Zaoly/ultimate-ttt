VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "井字棋游戏模式"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2415
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   2415
   StartUpPosition =   1  '所有者中心
   Begin VB.OptionButton Option4 
      Caption         =   "观看电脑内战"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2760
      Width           =   855
   End
   Begin VB.OptionButton Option3 
      Caption         =   "双人"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1935
   End
   Begin VB.OptionButton Option2 
      Caption         =   "单人，我后走（○）"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "单人，我先走（×）"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Value           =   -1  'True
      Width           =   1935
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    If Option1 Then
        GameMode = SoloFirst
    ElseIf Option2 Then
        GameMode = SoloSecond
    ElseIf Option3 Then
        GameMode = Duo
    ElseIf Option4 Then
        GameMode = Self
    End If
    Unload Form2
End Sub

Private Sub Command2_Click()
    If GameMode = 0 Then
        End
    Else
        IsCancelled = True
        Unload Form2
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Command2_Click
    End If
End Sub

Private Sub Form_Load()
    If GameMode <> 0 Then
        SwitchOrder
        Select Case GameMode
        Case SoloFirst
            Option1 = True
        Case SoloSecond
            Option2 = True
        Case Duo
            Option3 = True
        Case Self
            Option4 = True
        End Select
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Command2_Click
    End If
End Sub
