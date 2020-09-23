VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000001&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2625
   ClientLeft      =   10050
   ClientTop       =   0
   ClientWidth     =   4980
   ControlBox      =   0   'False
   ForeColor       =   &H0000FFFF&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer15 
      Interval        =   17000
      Left            =   4440
      Top             =   1560
   End
   Begin VB.Timer Timer14 
      Interval        =   16500
      Left            =   4440
      Top             =   1080
   End
   Begin VB.Timer Timer13 
      Interval        =   16000
      Left            =   4440
      Top             =   600
   End
   Begin VB.Timer Timer12 
      Interval        =   15500
      Left            =   4080
      Top             =   120
   End
   Begin VB.Timer Timer11 
      Interval        =   15000
      Left            =   3600
      Top             =   120
   End
   Begin VB.Timer Timer10 
      Interval        =   14500
      Left            =   3120
      Top             =   120
   End
   Begin VB.Timer Timer9 
      Interval        =   14000
      Left            =   2640
      Top             =   120
   End
   Begin VB.Timer Timer8 
      Interval        =   13500
      Left            =   2160
      Top             =   120
   End
   Begin VB.Timer Timer7 
      Interval        =   13000
      Left            =   1680
      Top             =   120
   End
   Begin VB.Timer Timer6 
      Interval        =   12500
      Left            =   1200
      Top             =   120
   End
   Begin VB.Timer Timer5 
      Interval        =   17500
      Left            =   2520
      Top             =   2040
   End
   Begin VB.Timer Timer4 
      Interval        =   10000
      Left            =   1920
      Top             =   2040
   End
   Begin VB.Timer Timer3 
      Interval        =   7500
      Left            =   1320
      Top             =   2040
   End
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   720
      Top             =   2040
   End
   Begin VB.Timer Timer1 
      Interval        =   2500
      Left            =   120
      Top             =   2040
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "Cathy"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Double Click to Stop"
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************************
'********************************************************************************
'** This code definitely needs help! While functional and fairly error proof,  **
'** it took a lot of typing, most I'm sure, unnecessary. This is my proverbial **
'** cry for help! I'm sure I could have set up arrays for the timers. Arrays   **
'** have troubled me as I seem to have a hard time grasping them. In addition  **
'** I would have liked to have had an input box so the user could type their   **
'** name instead of hard coding like I did. Finally, I selected a font color   **
'** I knew would not interfere with the system desktop color, thus rendering   **
'** it legible. While I was able to adapt the form BackColor property to the   **
'** system setting the font color property coding eluded me. If someone would  **
'** like to straighten this out, providing me the opportunity to learn, it     **
'** it very welcome and appreciated. Please excuse the generic object names    **
'** as I thought it might be easier for someone reading it to determine what   **
'** each thing was given the limited type of objects I used.                   **
'**                                                                            **
'** P.S. - I am fairly new to VB and come from the DOS world of batch files,   **
'**        Windows circa 3.1. I'm used to writing long "code" and can make it  **
'**        work but I know there are better ways to skin this cat. Nonetheless **
'**        the idea of creating my own programs, even this roughly coded one,  **
'**        is exciting and enjoyable. Thanks and feel free to modify and send  **
'**        me your revisions/ideas. I'll keep learning how to make this        **
'**        program better!                                  joefederico@email. **
'********************************************************************************
'********************************************************************************
Option Explicit
'
Dim nFirst As String
Dim nSecond As String
Dim nThird As String
Dim nFourth As String
Dim nFifth As String

Private Sub Form_Load()
    nFirst = Left$(Label1.Caption, 1)
    nSecond = Mid$(Label1.Caption, 1, 2)
    nThird = Mid$(Label1.Caption, 1, 3)
    nFourth = Mid$(Label1.Caption, 1, 4)
    nFifth = Mid$(Label1.Caption, 1, 5)
    Label1.Caption = nFirst
    Timer1.Enabled = True
    Timer2.Enabled = True
    Timer3.Enabled = True
    Timer4.Enabled = True
    Timer6.Enabled = True
    Timer7.Enabled = True
    Timer8.Enabled = True
    Timer9.Enabled = True
    Timer10.Enabled = True
    Timer11.Enabled = True
    Timer12.Enabled = True
    Timer13.Enabled = True
    Timer14.Enabled = True
    Timer15.Enabled = True
    Timer5.Enabled = True
End Sub

Private Sub Label1_DblClick()
    End
End Sub

Private Sub Timer1_Timer()
    Label1.Caption = nSecond
    Timer1.Enabled = False
End Sub

Private Sub Timer10_Timer()
    Label1.Visible = False
End Sub

Private Sub Timer11_Timer()
    Label1.Visible = True
End Sub

Private Sub Timer12_Timer()
    Label1.Visible = False
End Sub

Private Sub Timer13_Timer()
    Label1.Visible = True
End Sub

Private Sub Timer14_Timer()
    Label1.Visible = False
End Sub

Private Sub Timer15_Timer()
    Label1.Visible = True
End Sub

Private Sub Timer2_Timer()
    Label1.Caption = nThird
    Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
    Label1.Caption = nFourth
    Timer3.Enabled = False
End Sub

Private Sub Timer4_Timer()
    Label1.Caption = nFifth
    Timer4.Enabled = False
End Sub


Private Sub Timer5_Timer()
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    Timer4.Enabled = False
    Timer6.Enabled = False
    Timer7.Enabled = False
    Timer8.Enabled = False
    Timer9.Enabled = False
    Timer10.Enabled = False
    Timer11.Enabled = False
    Timer12.Enabled = False
    Timer13.Enabled = False
    Timer14.Enabled = False
    Timer15.Enabled = False
    Timer5.Enabled = False
    Label1.Caption = "Cathy"
    Call Form_Load
End Sub

Private Sub Timer6_Timer()
    Label1.Visible = False
End Sub

Private Sub Timer7_Timer()
    Label1.Visible = True
End Sub

Private Sub Timer8_Timer()
    Label1.Visible = False
End Sub

Private Sub Timer9_Timer()
    Label1.Visible = True
End Sub
