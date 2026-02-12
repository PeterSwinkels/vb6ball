VERSION 5.00
Begin VB.Form InterfaceWindow 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   209
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox BufferBox 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      FontTransparent =   0   'False
      Height          =   492
      Left            =   120
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   972
   End
End
Attribute VB_Name = "InterfaceWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This class contains this program's main interface window.
Option Explicit

'This procedure handles the user's keystrokes.
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrorTrap
   Select Case KeyCode
      Case vbKeyDown
         Paddle.XSpeed = 0
      Case vbKeyEscape
         Unload Me
      Case vbKeyLeft
         Paddle.XSpeed = -PADDLE_DEFAULT_SPEED
      Case vbKeyRight
         Paddle.XSpeed = PADDLE_DEFAULT_SPEED
   End Select
EndProcedure:
   Exit Sub

ErrorTrap:
   DisplayError
   Resume EndProcedure
End Sub

'This procedure initializes this window.
Private Sub Form_Load()
On Error GoTo ErrorTrap
   Me.Width = GAME_FIELD_WIDTH * Screen.TwipsPerPixelX
   Me.Width = Me.Width + (Me.Width - (Me.ScaleWidth * Screen.TwipsPerPixelX))

   Me.Height = GAME_FIELD_HEIGHT * Screen.TwipsPerPixelY
   Me.Height = Me.Height + (Me.Height - (Me.ScaleHeight * Screen.TwipsPerPixelY))
   
   BufferBox.Width = Me.ScaleWidth
   BufferBox.Height = Me.ScaleHeight
   
   Me.Caption = ProgramInformation()
EndProcedure:
   Exit Sub

ErrorTrap:
   DisplayError
   Resume EndProcedure
End Sub


