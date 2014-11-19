VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dhayalan Sivasuthan's Software Collection"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   8895
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Welcome.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   7455
   ScaleWidth      =   8895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Timer"
      Height          =   975
      Left            =   4800
      Picture         =   "Welcome.frx":1CCA
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Color Mixer"
      Height          =   975
      Left            =   3240
      Picture         =   "Welcome.frx":2594
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Explorer"
      Height          =   975
      Left            =   1680
      Picture         =   "Welcome.frx":325E
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Timer"
      Height          =   975
      Left            =   120
      Picture         =   "Welcome.frx":3F28
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculator"
      Height          =   1095
      Left            =   6720
      Picture         =   "Welcome.frx":47F2
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6240
      Width           =   2055
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7680
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   6240
      Top             =   120
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "If one of my softwares doesn't work or says that files are  missing, click here to intall library files."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   1680
      MouseIcon       =   "Welcome.frx":54BC
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Top             =   5880
      Width           =   6495
   End
   Begin VB.Image Image3 
      Height          =   1800
      Left            =   3720
      Picture         =   "Welcome.frx":57C6
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   2520
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   120
      Picture         =   "Welcome.frx":68F6
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1080
   End
   Begin VB.Image Image2 
      Height          =   960
      Left            =   5520
      Picture         =   "Welcome.frx":85C0
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1080
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Screen Saver 4"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      MouseIcon       =   "Welcome.frx":1A350
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   4800
      Width           =   3135
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Color Mixer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      MouseIcon       =   "Welcome.frx":1A65A
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   3360
      Width           =   3135
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Motherboard Quiz"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      MouseIcon       =   "Welcome.frx":1A964
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   5160
      Width           =   3135
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Age In Days"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      MouseIcon       =   "Welcome.frx":1AC6E
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   3720
      Width           =   3135
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Chemistry Quiz"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      MouseIcon       =   "Welcome.frx":1AF78
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   4440
      Width           =   3135
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Did You Know"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      MouseIcon       =   "Welcome.frx":1B282
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   4080
      Width           =   3135
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ID Reader"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      MouseIcon       =   "Welcome.frx":1B58C
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   2640
      Width           =   3135
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Book 3"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      MouseIcon       =   "Welcome.frx":1B896
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   4080
      Width           =   3135
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Media Player 2 (small)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      MouseIcon       =   "Welcome.frx":1BBA0
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   5520
      Width           =   3135
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Web Browser"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      MouseIcon       =   "Welcome.frx":1BEAA
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   3720
      Width           =   3135
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Text Editor"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      MouseIcon       =   "Welcome.frx":1C1B4
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   3000
      Width           =   3135
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Picture Viewer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      MouseIcon       =   "Welcome.frx":1C4BE
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   4440
      Width           =   3135
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "About the Author"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1320
      MouseIcon       =   "Welcome.frx":1C7C8
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   600
      Width           =   4095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Dhayalan Sivasuthan"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tamil On-screen Keyboard"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      MouseIcon       =   "Welcome.frx":1E492
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   5520
      Width           =   3135
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Calculator 3"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      MouseIcon       =   "Welcome.frx":1E79C
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   3360
      Width           =   3135
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Timer 4"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      MouseIcon       =   "Welcome.frx":1EAA6
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   2640
      Width           =   3135
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Explorer 6"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      MouseIcon       =   "Welcome.frx":1EDB0
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   3000
      Width           =   3135
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Media Player (large)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      MouseIcon       =   "Welcome.frx":1F0BA
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   5160
      Width           =   3135
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Converter 2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      MouseIcon       =   "Welcome.frx":1F3C4
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   4800
      Width           =   3135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Dhayalan Sivasuthan's self-made softwares"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   6495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
End
End If
If KeyCode = vbKeyEscape Then
End
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
End
End If
If KeyAscii = vbKeyEscape Then
End
End If
End Sub

Private Sub Form_Load()
'Me.Hide
'Form2.Show
MsgBox "If one of my softwares doesn't work properly, please install library files, hyperlink is provided at the bottom"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = vbRed
Label4.ForeColor = vbRed
Label5.ForeColor = vbRed
Label6.ForeColor = vbRed
Label7.ForeColor = vbRed
Label8.ForeColor = vbRed
Label9.FontUnderline = False
Label10.BackColor = &H8000000C
Label15.ForeColor = vbRed
Label16.ForeColor = vbRed
Label11.ForeColor = vbRed
Label12.ForeColor = vbRed
Label14.ForeColor = vbRed
Label13.ForeColor = vbRed
Label17.ForeColor = vbRed
Label18.ForeColor = vbRed
Label19.ForeColor = vbRed


Label22.ForeColor = vbRed
Label24.ForeColor = vbRed
Label25.ForeColor = vbRed
End Sub



Private Sub Form_Unload(Cancel As Integer)
Cancel = True
Form2.Show
Timer4.Enabled = True
End Sub



Private Sub Image3_Click()
Label10_Click
End Sub

Private Sub Label10_Click()
Me.Enabled = False
Form2.Show
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.BackColor = vbGreen
End Sub

Private Sub Label11_Click()
Shell "My Picture Viewer 2.exe", vbNormalFocus
End Sub

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label11.ForeColor = vbBlue
End Sub

Private Sub Label12_Click()
Shell "My Text Editor 2.exe", vbNormalFocus
End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label12.ForeColor = vbBlue
End Sub

Private Sub Label13_Click()
Shell "My Web Browser.exe", vbNormalFocus
End Sub

Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label13.ForeColor = vbBlue
End Sub

Private Sub Label14_Click()
Shell "My Media Player 3.exe", vbNormalFocus
End Sub

Private Sub Label14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label14.ForeColor = vbBlue
End Sub

Private Sub Label15_Click()
Shell "My Phone Book 3.exe", vbNormalFocus
End Sub

Private Sub Label15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label15.ForeColor = vbBlue
End Sub

Private Sub Label16_Click()
Shell "My ID Card Reader.exe", vbNormalFocus
End Sub

Private Sub Label16_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.ForeColor = vbBlue
End Sub

Private Sub Label17_Click()
Shell "Did You Know.exe", vbNormalFocus
End Sub

Private Sub Label17_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label17.ForeColor = vbBlue

End Sub

Private Sub Label22_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label22.ForeColor = vbBlue

End Sub

Private Sub Label24_Click()
Shell "My Color Mixer.exe", vbNormalFocus
End Sub

Private Sub Label25_Click()
Shell "My Screen Saver 4.exe", vbNormalFocus
End Sub

Private Sub Label3_Click()
Shell "My Unit Converter 2.exe", vbNormalFocus
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = vbBlue
End Sub

Private Sub Label4_Click()
Shell "My Media Player 2.exe", vbNormalFocus
End Sub


Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = vbBlue

End Sub

Private Sub Label5_Click()
Shell "My Explorer 6.exe", vbNormalFocus
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.ForeColor = vbBlue
End Sub

Private Sub Label6_Click()
Shell "My Timer 4.exe", vbNormalFocus
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = vbBlue
End Sub

Private Sub Label7_Click()
Shell "My Calculator 3.exe", vbNormalFocus
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.ForeColor = vbBlue
End Sub

Private Sub Label8_Click()
Shell "My Tamil Keyboard 2.exe", vbNormalFocus
End Sub



Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.ForeColor = vbBlue
End Sub

Private Sub Label9_Click()
Shell "libraryfiles.exe", vbNormalFocus
End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.FontUnderline = True
End Sub




Private Sub Timer2_Timer()
Label1.ForeColor = RGB(Rnd() * 800, Rnd() * 800, Rnd() * 800)
Label10.ForeColor = RGB(Rnd() * 500, Rnd() * 500, Rnd() * 800)
'Me.BackColor = RGB(Rnd() * 800, Rnd() * 800, Rnd() * 800)

End Sub

Private Sub Timer4_Timer()
If Me.Height < 1440 Then
End
Else
Me.Height = Me.Height - 25
End If
End Sub



Private Sub Label18_Click()
Shell "Chemistry Quiz.exe", vbNormalFocus
End Sub

Private Sub Label18_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label18.ForeColor = vbBlue
End Sub

Private Sub Label19_Click()
Shell "AgeInDays.exe", vbNormalFocus
End Sub

Private Sub Label19_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label19.ForeColor = vbBlue
End Sub

Private Sub Label20_Click()
Shell "My Web Browser.exe", vbNormalFocus
End Sub

Private Sub Label20_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label13.ForeColor = vbBlue
End Sub

Private Sub Label21_Click()
Shell "My Media Player 3.exe", vbNormalFocus
End Sub

Private Sub Label21_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label14.ForeColor = vbBlue
End Sub

Private Sub Label22_Click()
Shell "Motherboard Quiz.exe", vbNormalFocus
End Sub





Private Sub Label24_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label24.ForeColor = vbBlue
End Sub

Private Sub Label25_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label25.ForeColor = vbBlue
End Sub
