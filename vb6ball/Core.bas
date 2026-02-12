Attribute VB_Name = "CoreModule"
'This module contains this program's core procedures.
Option Explicit

'The Microsoft Windows API constants, functions and structures used by this program.

Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

Private Const SRCCOPY As Long = &HCC0020

Private Declare Function BitBlt Lib "Gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function QueryPerformanceCounter Lib "Kernel32.dll" (lpPerformanceCount As LARGE_INTEGER) As Long
Private Declare Function QueryPerformanceFrequency Lib "Kernel32.dll" (lpFrequency As LARGE_INTEGER) As Long


'The constants, structures, and variables used by this program.
Public Const GAME_FIELD_HEIGHT As Long = 480     'Defines the game field's height in pixels.
Public Const GAME_FIELD_WIDTH As Long = 640      'Defines the game field's width in pixels.
Public Const PADDLE_DEFAULT_SPEED As Long = 10   'Defines the paddle's default speed.
Private Const BALL_SIZE As Long = 10                  'Defines the ball's size in pixels.
Private Const MAXIMUM_SCORE As Long = 2147483647      'Defines the maximum possible score.
Private Const PADDLE_HEIGHT As Long = 16              'Defines the paddle's height in pixels.
Private Const PADDLE_WIDTH As Long = 80               'Defines the paddle's width in pixels.
Private Const SYMBOL_HEART_CHARACTER As Long = &HA9   'Defines the Symbol heart character.

'This structure defines the ball.
Private Type BallStr
   x As Long        'Defines the ball's horizontal position.
   y As Long        'Defines the ball's vertical position.
   XSpeed As Long   'Defines the ball's horizontal speed.
   YSpeed As Long   'Defines the ball's vertical speed.
End Type

'This structure defines the paddle.
Private Type PaddleStr
   x As Long        'Defines the paddle's horizontal position.
   y As Long        'Defines the paddle's vertical position.
   XSpeed As Long   'Defines the paddle's horizontal speed.
End Type

Public Paddle As PaddleStr               'Contains the paddle.
Private Ball As BallStr                   'Contains the ball.
Private DelayLength As LARGE_INTEGER      'Contains the most recent delay's length.
Private DelayStart As LARGE_INTEGER       'Contains the most recent delay's start.
Private Lives As Long                     'Contains the number of lives left.
Private Score As Long                     'Contains the player's score.
Private TicksPerSecond As LARGE_INTEGER   'Contains the number of ticks per second.

'This procedure controls the ball.
Private Sub ControlBall()
On Error GoTo ErrorTrap

   With Ball
      InterfaceWindow.BufferBox.FillColor = vbBlack
      InterfaceWindow.BufferBox.Circle (.x, .y), BALL_SIZE, vbBlack

      .x = .x + .XSpeed
      .y = .y + .YSpeed

      If .x < 0 Or .x >= GAME_FIELD_WIDTH Then
         .XSpeed = -.XSpeed
      End If

      If .y <= 0 Then
         .YSpeed = -.YSpeed
      ElseIf .y >= Paddle.y - BALL_SIZE Then
         If Ball.YSpeed > 0 Then
            If .x >= Paddle.x And .x < Paddle.x + PADDLE_WIDTH Then
               If Not Paddle.XSpeed = 0 Then
                  .XSpeed = Paddle.XSpeed
               End If
               .YSpeed = -.YSpeed
               Score = Score + 1
            ElseIf .y >= GAME_FIELD_HEIGHT Then
               .YSpeed = -.YSpeed
               Lives = Lives - 1
            End If
         End If
      End If
      
      InterfaceWindow.BufferBox.FillColor = vbGreen
      InterfaceWindow.BufferBox.Circle (.x, .y), BALL_SIZE, vbGreen
   End With

EndProcedure:
   Exit Sub

ErrorTrap:
   DisplayError
   Resume EndProcedure
End Sub

'This procedure controls the paddle.
Private Sub ControlPaddle()
On Error GoTo ErrorTrap

   With Paddle
      InterfaceWindow.BufferBox.Line (.x, .y)-Step(PADDLE_WIDTH, PADDLE_HEIGHT), vbBlack, BF

      .x = .x + .XSpeed

      If .x < 0 Or .x >= (GAME_FIELD_WIDTH - PADDLE_WIDTH) Then
         .XSpeed = 0
      End If
      
      InterfaceWindow.BufferBox.Line (.x, .y)-Step(PADDLE_WIDTH, PADDLE_HEIGHT), vbRed, BF
   End With

EndProcedure:
   Exit Sub

ErrorTrap:
   DisplayError
   Resume EndProcedure
End Sub


'This procedure displays any errors that occur.
Public Sub DisplayError()
Dim Description As String
Dim ErrorCode As Long

   Description = Err.Description
   ErrorCode = Err.Number

   On Error GoTo ErrorTrap
   MsgBox Description & vbCr & "Error code: " & CStr(ErrorCode), vbExclamation

EndProcedure:
   Exit Sub

EndProgram:
   End

ErrorTrap:
   Resume EndProgram
End Sub
'This procedure displays the status.
Private Sub DisplayStatus()
On Error GoTo ErrorTrap
   With InterfaceWindow.BufferBox
      .CurrentX = 0
      .CurrentY = 0
      .Font = "Comic Sans MS"
      .ForeColor = vbBlue
      InterfaceWindow.BufferBox.Print " Score: "; CStr(Score); " ";
   
      .CurrentX = GAME_FIELD_WIDTH * 0.9
      .CurrentY = 0
      .Font = "Symbol"
      .ForeColor = vbMagenta
      InterfaceWindow.BufferBox.Print String$(Lives, SYMBOL_HEART_CHARACTER); Space$(3);
   End With
EndProcedure:
   Exit Sub

ErrorTrap:
   DisplayError
   Resume EndProcedure
End Sub

'This procedure initializes this program.
Private Sub Initialize()
On Error GoTo ErrorTrap
   QueryPerformanceFrequency TicksPerSecond
   
   With Ball
      .x = 0
      .y = 0
      .XSpeed = 10
      .YSpeed = 10
   End With

   With Paddle
      .x = (GAME_FIELD_WIDTH / 2) - (PADDLE_WIDTH / 2)
      .XSpeed = 0
      .y = GAME_FIELD_HEIGHT - PADDLE_HEIGHT
   End With

   DelayLength.lowpart = TicksPerSecond.lowpart / 15
   DelayLength.highpart = TicksPerSecond.highpart / 15

   Lives = 3
   Score = 0
EndProcedure:
   Exit Sub

ErrorTrap:
   DisplayError
   Resume EndProcedure
End Sub

'This procedure is executed when this program is started.
Public Sub Main()
On Error GoTo ErrorTrap
Dim CurrentCounter As LARGE_INTEGER

   Initialize

   InterfaceWindow.Show

   QueryPerformanceCounter DelayStart
   Do While DoEvents() > 0
      QueryPerformanceCounter CurrentCounter
      If (CurrentCounter.highpart Xor &H80000000) >= (DelayStart.highpart Xor &H80000000) + DelayLength.highpart Then
         If (CurrentCounter.lowpart Xor &H80000000) >= (DelayStart.lowpart Xor &H80000000) + DelayLength.lowpart Then
            ControlBall
            ControlPaddle
            DisplayStatus
            With InterfaceWindow
               BitBlt .hDC, 0, 0, .ScaleWidth, .ScaleHeight, .BufferBox.hDC, 0, 0, SRCCOPY
            End With
            If Lives = 0 Then Exit Do
            If Score = MAXIMUM_SCORE Then Exit Do
            QueryPerformanceCounter DelayStart
         End If
      Else
         QueryPerformanceCounter DelayStart
      End If
   Loop

   If Lives = 0 Then
      MsgBox "Game over!", vbInformation
   End If

   If Score = MAXIMUM_SCORE Then
      MsgBox "You achieved the highest possible score!", vbInformation
   End If

EndProcedure:
   Exit Sub

ErrorTrap:
   DisplayError
   Resume EndProcedure
End Sub

'This procedure returns information about this program.
Public Function ProgramInformation() As String
On Error GoTo ErrorTrap
Dim Information As String

   With App
      Information = App.Title & " v" & App.Major & "." & App.Minor & App.Revision & " - by: " & App.CompanyName & ", " & App.LegalCopyright
   End With

EndProcedure:
   ProgramInformation = Information
   Exit Function

ErrorTrap:
   DisplayError
   Resume EndProcedure
End Function
