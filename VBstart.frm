VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "VB Start Menu Example, ph03n1x_2k@hotmail.com"
   ClientHeight    =   5055
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   Picture         =   "VBstart.frx":0000
   ScaleHeight     =   337
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   570
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6120
      Top             =   360
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   615
      Left            =   -120
      ScaleHeight     =   555
      ScaleWidth      =   8520
      TabIndex        =   0
      Top             =   4440
      Visible         =   0   'False
      Width           =   8580
      Begin VB.PictureBox Picture7 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   240
         Picture         =   "VBstart.frx":17C04
         ScaleHeight     =   255
         ScaleWidth      =   375
         TabIndex        =   7
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Start"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   1335
      End
      Begin VB.Frame Frame1 
         Height          =   495
         Left            =   7320
         TabIndex        =   2
         Top             =   0
         Width           =   975
         Begin VB.Label lbltime 
            Caption         =   "TIME"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   735
         End
      End
   End
   Begin VB.PictureBox docutab 
      Height          =   4455
      Left            =   2760
      ScaleHeight     =   4395
      ScaleWidth      =   2715
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   2775
      Begin VB.Label lblDoc 
         Alignment       =   2  'Center
         Caption         =   "Documents."
         Height          =   495
         Left            =   720
         TabIndex        =   9
         Top             =   1440
         Width           =   1455
      End
   End
   Begin VB.PictureBox favtab 
      Height          =   4455
      Left            =   2760
      ScaleHeight     =   4395
      ScaleWidth      =   2715
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   2775
      Begin VB.Label lblfav 
         Caption         =   "Favourites !!!"
         Height          =   615
         Left            =   600
         TabIndex        =   11
         Top             =   1440
         Width           =   2055
      End
   End
   Begin VB.PictureBox tabprog 
      Height          =   4455
      Left            =   2760
      ScaleHeight     =   4395
      ScaleWidth      =   2715
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   2775
      Begin VB.Label lblprog 
         Caption         =   "Programs"
         Height          =   255
         Left            =   840
         TabIndex        =   13
         Top             =   1440
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3450
      Left            =   0
      ScaleHeight     =   3420
      ScaleWidth      =   2745
      TabIndex        =   14
      Top             =   1080
      Visible         =   0   'False
      Width           =   2775
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   5775
         Left            =   0
         Picture         =   "VBstart.frx":17F78
         ScaleHeight     =   5775
         ScaleWidth      =   255
         TabIndex        =   21
         Top             =   -2280
         Width           =   255
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   240
         Picture         =   "VBstart.frx":18669
         ScaleHeight     =   615
         ScaleWidth      =   2535
         TabIndex        =   20
         Top             =   2760
         Width           =   2535
      End
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   240
         Picture         =   "VBstart.frx":18CD9
         ScaleHeight     =   615
         ScaleWidth      =   2535
         TabIndex        =   19
         Top             =   2160
         Width           =   2535
         Begin VB.Line Line1 
            X1              =   0
            X2              =   2880
            Y1              =   600
            Y2              =   600
         End
      End
      Begin VB.PictureBox Picture6 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   240
         Picture         =   "VBstart.frx":19285
         ScaleHeight     =   615
         ScaleWidth      =   2535
         TabIndex        =   18
         Top             =   1560
         Width           =   2535
      End
      Begin VB.PictureBox Picture8 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   240
         Picture         =   "VBstart.frx":19912
         ScaleHeight     =   495
         ScaleWidth      =   2655
         TabIndex        =   17
         Top             =   1080
         Width           =   2655
      End
      Begin VB.PictureBox Picture9 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   240
         Picture         =   "VBstart.frx":19F3B
         ScaleHeight     =   495
         ScaleWidth      =   2535
         TabIndex        =   16
         Top             =   600
         Width           =   2535
      End
      Begin VB.PictureBox Picture10 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   240
         Picture         =   "VBstart.frx":1A562
         ScaleHeight     =   615
         ScaleWidth      =   2535
         TabIndex        =   15
         Top             =   0
         Width           =   2535
      End
   End
   Begin VB.PictureBox settingstab 
      Height          =   4455
      Left            =   2760
      ScaleHeight     =   4395
      ScaleWidth      =   2715
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   2775
      Begin VB.Label settingstext 
         Alignment       =   2  'Center
         Caption         =   "Settings."
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   1440
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "mouse X,Y"
      Height          =   195
      Left            =   7560
      TabIndex        =   1
      Top             =   0
      Width           =   765
   End
   Begin VB.Menu menuabout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Note: the code I worked off for this was by Itay Sagui I left some of his code here including the help bit
' And I just made it into a start menu with some function's that actually work, be careful of the shutdown
' Button as it will shut down your computer (dont worry you get a message box asking to confirm though)


Private Sub Command2_Click()
' When you click the Command2 button picture one becomes visible.
    Picture1.Visible = True
    
End Sub

Private Sub Form_Click()

' When you click the form picture boxes hide
Picture1.Visible = False
Picture2.Visible = True
settingstab.Visible = False
tabprog.Visible = False
favtab.Visible = False
docutab.Visible = False


End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Tell the computer that H is going to be a integer (number)
Dim h As Integer
    h = Me.ScaleHeight - 8
    
    ' If the Y value is either more or equal to the H value then picture2 is visible.
    If (Y >= h) Then
        Picture2.Visible = True
    ' Other wise if the X value is less than the H value.
    ElseIf (X < h) Then
    ' Picture2 becomes visible.
        Picture2.Visible = False
    End If
    
    ' If you can see picture1 then picture2 will become visable, this is basicly so as
    ' You dont just get the top of the start menu with out the bar.
    
    If Picture1.Visible = True Then Picture2.Visible = True
    
    
    
    ' Label 1's text is the X and Y Co-ordinates or the mouse (remember your maths, X is always on the horizontal line, and Y the vertical)
    
    
    Label1 = X & "," & Y
End Sub

Private Sub Form_Resize()
 ' When you try and resize the form it automaticly goes back to the Height of 6780 and the width of 8460
  Me.Height = 5745
 Me.Width = 8460
 
End Sub


Private Sub menuabout_Click()
' Just displays this message when you click the about menu.
MsgBox "As you can see, creating a taskbar can be quite easy," & _
        " if one knows some tricks. You just create a picture box (or" & _
        " a frame), set their 'Visible' property to FALSE, change them" & _
        " to be left (or right) aligned. and set thier width. If the Form's" & _
        " MouseMove event, check if X is equal or lower than 8 (or what ever)" & _
        " you want your hot-zone to be, and if it is, set the picture box's" & _
        " to be visible." & vbNewLine & "  On the picture box that is used" & _
        " as the taskbar you can create anything you want."
End Sub

Private Sub Picture10_Click()
' Make the tab and the text visible.
lblprog.Visible = True
tabprog.Visible = True

' Hide the other tab's
docutab.Visible = False
favtab.Visible = False
settingstab.Visible = False

End Sub

Private Sub Picture4_Click()
' Pop's up a message box with Yes and No button's
MsgBox "Warning, this will shutdown your computer is it ok to continue ?", vbYesNo, "Warning!!"
' If the user clicks the yes button then pop up another message box
If vbYes Then MsgBox "This is your last chance save all your programs then click ok.", vbOKOnly, "Alert !!"
' as soon as OK is pressed the computer will shut down.
Shell "Rundll32 user,ExitWindows"


End Sub

Private Sub Picture5_Click()
' Show form 2
Form2.Show
End Sub

Private Sub Picture6_Click()
' When you click the settings button, the settings bit and the text inside get shown.
settingstab.Visible = True
settingstext.Visible = True

' Hide the other tab's
docutab.Visible = False
favtab.Visible = False
tabprog.Visible = False

End Sub

Private Sub Picture7_Click()
' Shows the start menu for people who like clicking the picture :P
    Picture1.Visible = True
End Sub

Private Sub Picture8_Click()
' Make the tab and the text visible.
docutab.Visible = True
lblDoc.Visible = True

' Hide the other tab's
settingstab.Visible = False
favtab.Visible = False
tabprog.Visible = False


End Sub

Private Sub Picture9_Click()
' Show the favourites stuff.
favtab.Visible = True
lblfav.Visible = True

' Hide the other tab's
docutab.Visible = False
settingstab.Visible = False
tabprog.Visible = False
End Sub

Private Sub Timer1_Timer()
' Label Time's caption = The time.
lbltime.Caption = Time
End Sub
