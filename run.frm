VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "VB Run."
   ClientHeight    =   1830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4950
   LinkTopic       =   "Form2"
   ScaleHeight     =   1830
   ScaleWidth      =   4950
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdrun 
      Caption         =   "Run"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   1320
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   0
      Picture         =   "run.frx":0000
      ScaleHeight     =   1335
      ScaleWidth      =   4935
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   840
         TabIndex        =   3
         Text            =   "Enter the program to be run here."
         Top             =   720
         Width           =   3615
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancel_Click()
'Hides this form.
Form2.Hide

End Sub

Private Sub cmdrun_Click()
' Just incase someone enters a invalid link.
On Error Resume Next

' Execute the path from text1 (in other words run the link thats in the text box)
' And set its window size to mazimized.
Shell Text1.Text, vbMaximizedFocus

End Sub

Private Sub Text1_Click()
' When someone clicks the text box the text box gets cleared.
Text1.Text = ""
End Sub
