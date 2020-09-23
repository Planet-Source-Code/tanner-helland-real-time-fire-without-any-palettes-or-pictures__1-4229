VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Tanner's Flame Creator"
   ClientHeight    =   1830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2280
   LinkTopic       =   "Form1"
   ScaleHeight     =   122
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   152
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1680
      Top             =   600
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ClipControls    =   0   'False
      FontTransparent =   0   'False
      Height          =   855
      Left            =   120
      MousePointer    =   11  'Hourglass
      ScaleHeight     =   53
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   53
      TabIndex        =   0
      ToolTipText     =   "Here's your flame!"
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "***"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Real-time flames by Tanner Helland

'This great program creates real-time flames using a simple little
'method I thought up in math one day.  Basically, it runs a loop
'from the bottom of the picture to the top of the picture, drawing
'pixels of varying brightness depending on their height and a
'randomly determined amount.  It takes a little while to get started,
'but once it gets going the flames move pretty realistically.  Also,
'the way it gets the different colors is pretty snazzy - I hate
'trying to use palettes in VB because it's slow and obnoxious and
'tons of work, so I devised this method: double the red value, keep
'the green value the same, and halve the blue value.  This results
'in only red, orange, and yellow shades of pixels.  Cool, huh?  If
'you like what you see, post any comments and/or suggestions at
'planet source code.

'Feel free to use this code however you want, but please give
'me some credit and let me know how you use it.  E-mail me with
'comments or questions at tannerhelland@hotmail.com.

'Also, if you like graphic-based programming, look for some
'of my other projects at www.planet-source-code.com.  Also, feel
'free to e-mail me with any graphic-related (or VB in general)
'questions you may have.

'Read the enclosed game.txt file for info on our current game
'production 'Realms of Time.'

Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
'The array for the separate flame pixels
Private FlameArray() As Byte
'counts the frame
Private Frame As Integer
'this number is used to fade the colors to their different
'values, as well as determine the intensity (2nd #)
Const temp = 256 / 50

Private Sub Form_Load()
'sets the array to the correct size
ReDim FlameArray(0 To 50, 0 To 50) As Byte
'turns the bottom row of flames "on"
For x = 0 To 50
For y = 46 To 50
FlameArray(x, y) = 50
Next y
Next x
On Error Resume Next

Kill "Flames.lst"
End Sub

Private Sub Timer1_Timer()
On Error Resume Next

Static x As Integer
Static y As Integer
Static Color As Integer
Static temp2 As Byte
'runs the loop for the y-axis
For y = 50 To 4 Step -1
'runs the loop for the x-axis
For x = 0 To 50
'set the random degree of cooldown
FlameArray(x, y) = FlameArray(x, y) - Int(Rnd * 3)
'set a new random number for the movement
temp2 = Int(Rnd * 3)
'move the pixel in the array
FlameArray(x, y - temp2) = FlameArray(x, y)
'get the color based on the "heat" value
Color = (Int(FlameArray(x, y) * temp))
'draw the pixel
SetPixel Picture1.hDC, x + (Rnd * 2), y, RGB(Color + Color, Color, Color / 2)
'end x loop
Next x
'end y loop
Next y

'make the bottom 4 rows hot again
For x = 0 To 50
For y = 46 To 50
FlameArray(x, y) = 50
Next y
Next x
Picture1.Refresh

'If you want, you could build a despeckle loop right here that may
'make the flames more realistic.  Maybe someday I'll get around to
'it, but in the meantime feel free to edit this as you like and
'send me the updates.

Exit Sub
'this is to save the frames as bitmaps

Static PicName As String
    Frame = Frame + 1
    Label1.Caption = Frame
If Frame > 20 Then
PicName = "picflame" & (Frame - 15) & ".bmp"
SavePicture Picture1.Image, PicName
Open "Flames.lst" For Append As #1
Print #1, PicName
Close #1
End If

End Sub
