VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E6FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "无敌井字棋 (当前水平: 无敌)"
   ClientHeight    =   5295
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   5295
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   353
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   353
   StartUpPosition =   3  '窗口缺省
   Begin VB.Label LabelNum 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   3480
      TabIndex        =   8
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label LabelNum 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   1920
      TabIndex        =   7
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label LabelNum 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   360
      TabIndex        =   6
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label LabelNum 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   3480
      TabIndex        =   5
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label LabelNum 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1920
      TabIndex        =   4
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label LabelNum 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   360
      TabIndex        =   3
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label LabelNum 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3480
      TabIndex        =   2
      Top             =   360
      Width           =   255
   End
   Begin VB.Label LabelNum 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   255
   End
   Begin VB.Label LabelNum 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   255
   End
   Begin VB.Image ImageSquare 
      Height          =   1455
      Index           =   8
      Left            =   3480
      Picture         =   "无敌井字棋.frx":0000
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Image ImageSquare 
      Height          =   1455
      Index           =   7
      Left            =   1920
      Picture         =   "无敌井字棋.frx":6EE6
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Image ImageSquare 
      Height          =   1455
      Index           =   6
      Left            =   360
      Picture         =   "无敌井字棋.frx":DDCC
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Image ImageSquare 
      Height          =   1455
      Index           =   5
      Left            =   3480
      Picture         =   "无敌井字棋.frx":14CB2
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Image ImageSquare 
      Height          =   1455
      Index           =   4
      Left            =   1920
      Picture         =   "无敌井字棋.frx":1BB98
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Image ImageSquare 
      Height          =   1455
      Index           =   3
      Left            =   360
      Picture         =   "无敌井字棋.frx":22A7E
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Image ImageSquare 
      Height          =   1455
      Index           =   2
      Left            =   3480
      Picture         =   "无敌井字棋.frx":29964
      Top             =   360
      Width           =   1455
   End
   Begin VB.Image ImageSquare 
      Height          =   1455
      Index           =   1
      Left            =   1920
      Picture         =   "无敌井字棋.frx":3084A
      Top             =   360
      Width           =   1455
   End
   Begin VB.Image ImageSquare 
      Height          =   1455
      Index           =   0
      Left            =   360
      Picture         =   "无敌井字棋.frx":37730
      Top             =   360
      Width           =   1455
   End
   Begin VB.Menu MenuGame 
      Caption         =   "游戏(&G)"
      NegotiatePosition=   1  'Left
      Begin VB.Menu MenuComputerLevel 
         Caption         =   "电脑水平(&L)..."
         Begin VB.Menu MenuEasy 
            Caption         =   "初级(&E)"
         End
         Begin VB.Menu MenuNormal 
            Caption         =   "中级(&N)"
         End
         Begin VB.Menu MenuHard 
            Caption         =   "高级(&H)"
         End
         Begin VB.Menu MenuUltimate 
            Caption         =   "无敌!(&U)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu MenuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu Undo 
         Caption         =   "悔棋？想多了吧..."
         Enabled         =   0   'False
      End
      Begin VB.Menu MenuReplay 
         Caption         =   "重玩(&R)"
         Shortcut        =   ^R
      End
      Begin VB.Menu MenuDash2 
         Caption         =   "-"
      End
      Begin VB.Menu MenuQuit 
         Caption         =   "退出(&Q)"
      End
   End
   Begin VB.Menu MenuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu MenuHowToPlay 
         Caption         =   "玩法(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu MenuAbout 
         Caption         =   "关于(&A)"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long

Private Enum PieceType
    NoPiece
    Piece1
    Piece2
End Enum

Private Enum ResultType
    ResUnknown
    ResWin
    ResDraw
    ResLose
End Enum

Private Const IDC_HAND As Long = 32649& ' When mouse over square, mouse shape turns into a hand

Dim TurnNo As Integer, Moves(8) As Integer, Pieces(8) As PieceType, Board(2, 2) As PieceType
Dim IsMouseOver(8) As Boolean, IsMouseDisabled As Boolean, WinLoseReversed As Boolean
Dim hHandCursor As Long

Private Function RandInt(Min As Long, Max As Long) As Long
    Randomize
    RandInt = Int(Rnd * (Max - Min + 1)) + Min
End Function

' Position number & array conversion:

' Array: (R = row, C = column; array[R][C])

'  R\C  0 1 2

'   0   A B C
'   1   D E F
'   2   G H I

' Number:
'  A * 1 + B * 3 + C * 9 + ... + I * 6561

Private Function BoardToNum() As Integer ' Transform position array to position number
    Dim Row As Integer, Column As Integer, Result As Integer
    For Row = 0 To 2
        For Column = 0 To 2
            Result = Result + Board(Row, Column) * 3 ^ (Row * 3 + Column)
        Next
    Next
    BoardToNum = Result
End Function

Private Function MoveToBoard(Move As Integer) As Integer ' Move on current board
    If (TurnNo And 1) = 0 Then
        MoveToBoard = BoardToNum + Piece1 * 3 ^ Move
    ElseIf (TurnNo And 1) = 1 Then
        MoveToBoard = BoardToNum + Piece2 * 3 ^ Move
    End If
End Function

Private Sub ChangeImageSquare(Index As Integer) ' Redraw square
    Dim FileName As String
    FileName = "square"
    Select Case Pieces(Index)
    Case NoPiece
        If IsMouseOver(Index) Then
            FileName = FileName & "Over"
            If TurnNo Mod 2 = 0 Then
                FileName = FileName & "X"
            ElseIf TurnNo Mod 2 = 1 Then
                FileName = FileName & "O"
            End If
        End If
    Case Piece1
        FileName = FileName & "X"
    Case Piece2
        FileName = FileName & "O"
    End Select
    FileName = FileName & ".bmp"
    ImageSquare(Index).Picture = LoadPicture(FileName)
End Sub

Private Function JudgePosition() As ResultType
    Dim Does1Win As Boolean, Does2Win As Boolean
    Does1Win = _
        Board(0, 0) = Board(0, 1) And Board(0, 1) = Board(0, 2) And Board(0, 2) = Piece1 Or _
        Board(1, 0) = Board(1, 1) And Board(1, 1) = Board(1, 2) And Board(1, 2) = Piece1 Or _
        Board(2, 0) = Board(2, 1) And Board(2, 1) = Board(2, 2) And Board(2, 2) = Piece1 Or _
        Board(0, 0) = Board(1, 0) And Board(1, 0) = Board(2, 0) And Board(2, 0) = Piece1 Or _
        Board(0, 1) = Board(1, 1) And Board(1, 1) = Board(2, 1) And Board(2, 1) = Piece1 Or _
        Board(0, 2) = Board(1, 2) And Board(1, 2) = Board(2, 2) And Board(2, 2) = Piece1 Or _
        Board(0, 0) = Board(1, 1) And Board(1, 1) = Board(2, 2) And Board(2, 2) = Piece1 Or _
        Board(0, 2) = Board(1, 1) And Board(1, 1) = Board(2, 0) And Board(2, 0) = Piece1
    ' Including:
    
    '  111    000    000    100    010    001    100    001
    '  000    111    000    100    010    001    010    010
    '  000    000    111    100    010    001    001    100
    
    Does2Win = _
        Board(0, 0) = Board(0, 1) And Board(0, 1) = Board(0, 2) And Board(0, 2) = Piece2 Or _
        Board(1, 0) = Board(1, 1) And Board(1, 1) = Board(1, 2) And Board(1, 2) = Piece2 Or _
        Board(2, 0) = Board(2, 1) And Board(2, 1) = Board(2, 2) And Board(2, 2) = Piece2 Or _
        Board(0, 0) = Board(1, 0) And Board(1, 0) = Board(2, 0) And Board(2, 0) = Piece2 Or _
        Board(0, 1) = Board(1, 1) And Board(1, 1) = Board(2, 1) And Board(2, 1) = Piece2 Or _
        Board(0, 2) = Board(1, 2) And Board(1, 2) = Board(2, 2) And Board(2, 2) = Piece2 Or _
        Board(0, 0) = Board(1, 1) And Board(1, 1) = Board(2, 2) And Board(2, 2) = Piece2 Or _
        Board(0, 2) = Board(1, 1) And Board(1, 1) = Board(2, 0) And Board(2, 0) = Piece2
    If Does1Win And Not Does2Win Then     ' First player wins
        If TurnNo Mod 2 = 0 Then     ' Second player's turn
            JudgePosition = ResLose
        ElseIf TurnNo Mod 2 = 1 Then ' First player's turn
            JudgePosition = ResWin
        End If
    ElseIf Does2Win And Not Does1Win Then ' Second player wins
        If TurnNo Mod 2 = 0 Then     ' Second player's turn
            JudgePosition = ResWin
        ElseIf TurnNo Mod 2 = 1 Then ' First player's turn
            JudgePosition = ResLose
        End If
    ElseIf Does1Win And Does2Win Or _
        Board(0, 0) <> NoPiece And Board(0, 1) <> NoPiece And Board(0, 2) <> NoPiece And _
        Board(1, 0) <> NoPiece And Board(1, 1) <> NoPiece And Board(1, 2) <> NoPiece And _
        Board(2, 0) <> NoPiece And Board(2, 1) <> NoPiece And Board(2, 2) <> NoPiece _
    Then ' Board is full
        JudgePosition = ResDraw
    Else ' Game not over yet
        JudgePosition = ResUnknown
    End If
End Function

Private Sub ClearBoard()
    Dim Index As Integer
    For Index = 0 To 8
        If Pieces(Index) <> NoPiece Or IsMouseOver(Index) Then ' Clear board, where "Index" is of piece
            Pieces(Index) = NoPiece
            Board(Index \ 3, Index Mod 3) = NoPiece
            IsMouseOver(Index) = False
            ChangeImageSquare Index
        End If
        Moves(Index) = 0 ' Clear moves, where "Index" is of move
    Next
    TurnNo = 0
    If GameMode <> Self Then
        AutoMove
    End If
End Sub

Private Sub AutoMove()
    Const MaskData = "TicTacToeMoveTable ThisIsDevelopedByZaoly PleaseDoNotModifyMe SoAsNotToInterfereWithAutoMover ThankYouForCooperation Love Zaoly "
    Dim BoardNum As Integer, ByteValue As Byte, LargestResultValue As Integer, ResultValue As Integer, Move As Integer
    Dim BestMoves() As Integer, BestMoveCount As Integer, BestMoveIndex As Integer
    If _
        Len(Dir("tic-tac-toe.mvtbl")) = 0 Or _
        Len(Dir("tic-tac-toe-rev.mvtbl")) = 0 _
    Then ' Check if file exists
        End
    End If
    If Not ( _
        GameMode = SoloFirst And TurnNo Mod 2 = 1 Or _
        GameMode = SoloSecond And TurnNo Mod 2 = 0 Or _
        GameMode = Self _
    ) Then ' Check game mode
        Exit Sub
    End If
    IsMouseDisabled = True
    DoEvents
    If WinLoseReversed Then ' Surprise mode?
        Open "tic-tac-toe-rev.mvtbl" For Binary As #1 ' Reverse it!
    Else
        Open "tic-tac-toe.mvtbl" For Binary As #1
    End If
    If LOF(1) <> 19683 Then ' Check file size
        End
    End If
    BoardNum = BoardToNum
    Get #1, BoardNum + 1, ByteValue
    ByteValue = ByteValue Xor CByte(Asc(Mid(MaskData, (BoardNum And 127) + 1, 1))) ' Decode
    If ByteValue <= 127 Then
        LargestResultValue = ByteValue
    ElseIf ByteValue >= 128 Then
        LargestResultValue = ByteValue - 256
    End If
    If LargestResultValue > 0 Then
        LargestResultValue = -2 - LargestResultValue
    ElseIf LargestResultValue < 0 Then
        LargestResultValue = -LargestResultValue
    End If
    If _
        MenuEasy.Checked Or _
        MenuNormal.Checked And LargestResultValue > -127 And LargestResultValue < 127 Or _
        MenuHard.Checked And LargestResultValue > -125 And LargestResultValue < 125 _
    Then ' Easy mode: idiot
         ' Normal mode: only 1~2 steps considered
         ' Hard mode: only 2~3 steps considered
         ' Ultimate mode: eh hey hey hey~ (BEAT ME? NO WAY!)
        LargestResultValue = 0
    End If
    For Move = 0 To 8
        If Pieces(Move) = NoPiece Then
            BoardNum = MoveToBoard(Move)
            Get #1, BoardNum + 1, ByteValue
            ByteValue = ByteValue Xor CByte(Asc(Mid(MaskData, (BoardNum And 127) + 1, 1))) ' Decode
            If ByteValue <= 127 Then
                ResultValue = ByteValue
            ElseIf ByteValue >= 128 Then
                ResultValue = ByteValue - 256
            End If
            If _
                MenuEasy.Checked Or _
                MenuNormal.Checked And ResultValue > -127 And LargestResultValue < 127 Or _
                MenuHard.Checked And ResultValue > -125 And LargestResultValue < 125 _
            Then
                ResultValue = 0
            End If
            If ResultValue = LargestResultValue Then
                ReDim Preserve BestMoves(BestMoveCount)
                BestMoves(BestMoveCount) = Move
                BestMoveCount = BestMoveCount + 1
            End If
        End If
    Next
    Close #1
    BestMoveIndex = RandInt(0, BestMoveCount - 1)
    IsMouseDisabled = False
    ImageSquare_MouseDown BestMoves(BestMoveIndex), vbLeftButton, 0, ImageSquare(BestMoves(BestMoveIndex)).Left, ImageSquare(BestMoves(BestMoveIndex)).Top
End Sub

Private Sub HighlightPieces(MyPiece As PieceType, Square1 As Integer, Square2 As Integer, Square3 As Integer)
    If MyPiece = Piece1 Then
        ImageSquare(Square1).Picture = LoadPicture("squareX!.bmp")
        ImageSquare(Square2).Picture = LoadPicture("squareX!.bmp")
        ImageSquare(Square3).Picture = LoadPicture("squareX!.bmp")
    ElseIf MyPiece = Piece2 Then
        ImageSquare(Square1).Picture = LoadPicture("squareO!.bmp")
        ImageSquare(Square2).Picture = LoadPicture("squareO!.bmp")
        ImageSquare(Square3).Picture = LoadPicture("squareO!.bmp")
    End If
End Sub

Private Sub ShowWin()
    Select Case True
    Case Board(0, 0) = Board(0, 1) And Board(0, 1) = Board(0, 2) And Board(0, 2) = Piece1
        HighlightPieces Piece1, 0, 1, 2
    Case Board(1, 0) = Board(1, 1) And Board(1, 1) = Board(1, 2) And Board(1, 2) = Piece1
        HighlightPieces Piece1, 3, 4, 5
    Case Board(2, 0) = Board(2, 1) And Board(2, 1) = Board(2, 2) And Board(2, 2) = Piece1
        HighlightPieces Piece1, 6, 7, 8
    Case Board(0, 0) = Board(1, 0) And Board(1, 0) = Board(2, 0) And Board(2, 0) = Piece1
        HighlightPieces Piece1, 0, 3, 6
    Case Board(0, 1) = Board(1, 1) And Board(1, 1) = Board(2, 1) And Board(2, 1) = Piece1
        HighlightPieces Piece1, 1, 4, 7
    Case Board(0, 2) = Board(1, 2) And Board(1, 2) = Board(2, 2) And Board(2, 2) = Piece1
        HighlightPieces Piece1, 2, 5, 8
    Case Board(0, 0) = Board(1, 1) And Board(1, 1) = Board(2, 2) And Board(2, 2) = Piece1
        HighlightPieces Piece1, 0, 4, 8
    Case Board(0, 2) = Board(1, 1) And Board(1, 1) = Board(2, 0) And Board(2, 0) = Piece1
        HighlightPieces Piece1, 2, 4, 6
    Case Board(0, 0) = Board(0, 1) And Board(0, 1) = Board(0, 2) And Board(0, 2) = Piece2
        HighlightPieces Piece2, 0, 1, 2
    Case Board(1, 0) = Board(1, 1) And Board(1, 1) = Board(1, 2) And Board(1, 2) = Piece2
        HighlightPieces Piece2, 3, 4, 5
    Case Board(2, 0) = Board(2, 1) And Board(2, 1) = Board(2, 2) And Board(2, 2) = Piece2
        HighlightPieces Piece2, 6, 7, 8
    Case Board(0, 0) = Board(1, 0) And Board(1, 0) = Board(2, 0) And Board(2, 0) = Piece2
        HighlightPieces Piece2, 0, 3, 6
    Case Board(0, 1) = Board(1, 1) And Board(1, 1) = Board(2, 1) And Board(2, 1) = Piece2
        HighlightPieces Piece2, 1, 4, 7
    Case Board(0, 2) = Board(1, 2) And Board(1, 2) = Board(2, 2) And Board(2, 2) = Piece2
        HighlightPieces Piece2, 2, 5, 8
    Case Board(0, 0) = Board(1, 1) And Board(1, 1) = Board(2, 2) And Board(2, 2) = Piece2
        HighlightPieces Piece2, 0, 4, 8
    Case Board(0, 2) = Board(1, 1) And Board(1, 1) = Board(2, 0) And Board(2, 0) = Piece2
        HighlightPieces Piece2, 2, 4, 6
    End Select
End Sub

Private Sub InitBoard()
    Select Case True
    Case MenuEasy.Checked
        Form1.Caption = "无敌井字棋 (当前水平: 初级)"
    Case MenuNormal.Checked
        Form1.Caption = "无敌井字棋 (当前水平: 中级)"
    Case MenuHard.Checked
        Form1.Caption = "无敌井字棋 (当前水平: 高级)"
    Case MenuUltimate.Checked
        Form1.Caption = "无敌井字棋 (当前水平: 无敌)"
    End Select
    ClearBoard
    If GameMode = Self Then
        MsgBox "你可以决定第一步棋。", vbInformation
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 49 And KeyAscii <= 57 Then
        ImageSquare_MouseDown KeyAscii - 49, vbLeftButton, 0, ImageSquare(KeyAscii - 49).Left, ImageSquare(KeyAscii - 49).Top
    End If
End Sub

Private Sub Form_Load()
    hHandCursor = LoadCursor(0&, IDC_HAND) ' Load hand mouse icon
    Form1.Show
    Form2.Show vbModal
    InitBoard
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Index As Integer
    For Index = 0 To 8
        If IsMouseOver(Index) Then
            IsMouseOver(Index) = False
            ChangeImageSquare Index
        End If
    Next
End Sub

Private Sub ImageSquare_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim GameResult As ResultType, PromptContent As String
    If Button <> vbLeftButton Or IsMouseDisabled Or Pieces(Index) <> NoPiece Then
        Exit Sub
    End If
    If TurnNo Mod 2 = 0 Then
        Pieces(Index) = Piece1
        Board(Index \ 3, Index Mod 3) = Piece1
    ElseIf TurnNo Mod 2 = 1 Then
        Pieces(Index) = Piece2
        Board(Index \ 3, Index Mod 3) = Piece2
    End If
    IsMouseOver(Index) = False
    ChangeImageSquare Index
    Moves(TurnNo) = Index
    TurnNo = TurnNo + 1
    GameResult = JudgePosition
    If GameResult = ResWin Then
        ShowWin
        If TurnNo Mod 2 = 0 Xor WinLoseReversed Then
            Form1.BackColor = RGB(0, 96, 162)
            If GameMode = SoloFirst Then
                PromptContent = "很抱歉，你输了！"
            ElseIf GameMode = SoloSecond Then
                PromptContent = "恭喜你，你赢了！"
            Else
                PromptContent = "游戏结束，恭喜 ○ 方（后手）胜！"
            End If
            If MsgBox(PromptContent & vbCrLf & _
                      vbCrLf & _
                      "再来一局吗？", _
            vbInformation + vbYesNo) = vbNo Then
                End
            End If
        ElseIf TurnNo Mod 2 = 1 Xor WinLoseReversed Then
            Form1.BackColor = RGB(191, 95, 29)
            If GameMode = SoloFirst Then
                PromptContent = "恭喜你，你赢了！"
            ElseIf GameMode = SoloSecond Then
                PromptContent = "很抱歉，你输了！"
            Else
                PromptContent = "游戏结束，恭喜 × 方（先手）胜！"
            End If
            If MsgBox(PromptContent & vbCrLf & _
                      vbCrLf & _
                      "再来一局吗？", _
            vbInformation + vbYesNo) = vbNo Then
                End
            End If
        End If
        If WinLoseReversed Then
            Form1.BackColor = RGB(255, 230, 255)
        Else
            Form1.BackColor = RGB(255, 255, 230)
        End If
        SwitchOrder
        ClearBoard
    ElseIf GameResult = ResDraw Then
        If MsgBox("游戏结束，和棋！" & vbCrLf & _
                  vbCrLf & _
                  "再来一局吗？", _
        vbInformation + vbYesNo) = vbNo Then
            End
        End If
        SwitchOrder
        ClearBoard
    Else
        AutoMove
    End If
End Sub

Private Sub ImageSquare_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim CurrentIndex As Integer
    If Pieces(Index) = NoPiece And Not IsMouseDisabled Then
        SetCursor hHandCursor
    End If
    For CurrentIndex = 0 To 8
        If IsMouseOver(CurrentIndex) <> (CurrentIndex = Index And Pieces(Index) = NoPiece) Then
            IsMouseOver(CurrentIndex) = CurrentIndex = Index And Pieces(Index) = NoPiece
            If Not IsMouseDisabled Then
                ChangeImageSquare CurrentIndex
            End If
        End If
    Next
End Sub

Private Sub MenuAbout_Click()
    Static ClickTimes As Long
    MsgBox "无敌井字棋" & vbCrLf & _
           vbCrLf & _
           "开发者: 赵迪 (Zaoly)" & vbCrLf & _
           "QQ: 2298511336", _
    vbInformation
    ClickTimes = ClickTimes + 1
    If ClickTimes = 3 Then
        MsgBox "恭喜你，你发现了本游戏的“彩蛋”——胜负反转模式。" & vbCrLf & _
               "此模式下，有三子连成一排的一方反而为输。" & vbCrLf & _
               vbCrLf & _
               "祝您游戏愉快！（若要退出“胜负反转”模式，请重启游戏。）", _
        vbInformation
        WinLoseReversed = True
        Form1.BackColor = RGB(255, 230, 255)
        InitBoard
    End If
End Sub

Private Sub MenuEasy_Click()
    MenuEasy.Checked = True
    MenuNormal.Checked = False
    MenuHard.Checked = False
    MenuUltimate.Checked = False
    InitBoard
End Sub

Private Sub MenuHard_Click()
    MenuEasy.Checked = False
    MenuNormal.Checked = False
    MenuHard.Checked = True
    MenuUltimate.Checked = False
    InitBoard
End Sub

Private Sub MenuHowToPlay_Click()
    MsgBox "单击方格即可下棋，× 先手，○ 后手。" & vbCrLf & _
           "电脑难度无敌，绝不会输棋。" & vbCrLf & _
           vbCrLf & _
           "注：按数字键盘也可下棋哦！", _
    vbInformation
End Sub

Private Sub MenuNormal_Click()
    MenuEasy.Checked = False
    MenuNormal.Checked = True
    MenuHard.Checked = False
    MenuUltimate.Checked = False
    InitBoard
End Sub

Private Sub MenuQuit_Click()
    End
End Sub

Private Sub MenuReplay_Click()
    Form2.Show vbModal
    If IsCancelled Then
        IsCancelled = False
    Else
        InitBoard
    End If
End Sub

Private Sub MenuUltimate_Click()
    MenuEasy.Checked = False
    MenuNormal.Checked = False
    MenuHard.Checked = False
    MenuUltimate.Checked = True
    InitBoard
End Sub
