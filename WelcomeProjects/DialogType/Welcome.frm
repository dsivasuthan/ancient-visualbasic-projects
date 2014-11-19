VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dhayalan Sivasuthan's Software Collection"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   8145
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
   ScaleHeight     =   5520
   ScaleWidth      =   8145
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command17 
      Caption         =   "About Author"
      Height          =   1815
      Left            =   6480
      Picture         =   "Welcome.frx":1CCA
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton Command16 
      Caption         =   "General Quiz"
      Height          =   855
      Left            =   8040
      Picture         =   "Welcome.frx":3994
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   3480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Calendar"
      Height          =   855
      Left            =   6480
      Picture         =   "Welcome.frx":465E
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Unit Converter"
      Height          =   855
      Left            =   4920
      Picture         =   "Welcome.frx":6328
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Motherboad Quiz"
      Height          =   855
      Left            =   240
      Picture         =   "Welcome.frx":6BF2
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Tamil Keyboard"
      Height          =   855
      Left            =   1800
      Picture         =   "Welcome.frx":74BC
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Picture Viewer"
      Height          =   855
      Left            =   3360
      Picture         =   "Welcome.frx":8186
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Phone Book"
      Height          =   855
      Left            =   3360
      Picture         =   "Welcome.frx":8E50
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Age In Days"
      Height          =   855
      Left            =   1800
      Picture         =   "Welcome.frx":9B1A
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Text Editor"
      Height          =   855
      Left            =   240
      Picture         =   "Welcome.frx":A7E4
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Command10 
      Caption         =   "SL ID Card Reader"
      Height          =   855
      Left            =   4920
      Picture         =   "Welcome.frx":B4AE
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Media Player (ii)"
      Height          =   855
      Left            =   3360
      Picture         =   "Welcome.frx":C378
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Media Player (i)"
      Height          =   855
      Left            =   1800
      Picture         =   "Welcome.frx":D042
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Chemistry Quiz"
      Height          =   855
      Left            =   240
      Picture         =   "Welcome.frx":DD0C
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Descriptions"
      Height          =   855
      Left            =   4920
      Picture         =   "Welcome.frx":E016
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Web Browser"
      Height          =   855
      Left            =   4920
      Picture         =   "Welcome.frx":ECE0
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Color Mixer"
      Height          =   855
      Left            =   3360
      Picture         =   "Welcome.frx":F9AA
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Explorer"
      Height          =   855
      Left            =   1800
      Picture         =   "Welcome.frx":10674
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Timer"
      Height          =   855
      Left            =   6480
      Picture         =   "Welcome.frx":1133E
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculator"
      Height          =   855
      Left            =   240
      Picture         =   "Welcome.frx":11C08
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3480
      Width           =   1455
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
      Caption         =   "If any of my softwares doesn't work or says that files are  missing, click here to intall library files."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   495
      Left            =   240
      MouseIcon       =   "Welcome.frx":128D2
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Top             =   5640
      Width           =   7695
   End
   Begin VB.Image Image3 
      Height          =   1800
      Left            =   5640
      Picture         =   "Welcome.frx":12BDC
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   2520
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   120
      Picture         =   "Welcome.frx":13D0C
      Stretch         =   -1  'True
      Top             =   120
      Width           =   960
   End
   Begin VB.Image Image2 
      Height          =   960
      Left            =   6960
      Picture         =   "Welcome.frx":159D6
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
      Left            =   3360
      MouseIcon       =   "Welcome.frx":27766
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   8280
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
      Left            =   3360
      MouseIcon       =   "Welcome.frx":27A70
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   6840
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
      Left            =   120
      MouseIcon       =   "Welcome.frx":27D7A
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   8640
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
      Left            =   120
      MouseIcon       =   "Welcome.frx":28084
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   7200
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
      Left            =   3360
      MouseIcon       =   "Welcome.frx":2838E
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   7920
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
      Left            =   3360
      MouseIcon       =   "Welcome.frx":28698
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   7560
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
      Left            =   120
      MouseIcon       =   "Welcome.frx":289A2
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   6120
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
      Left            =   120
      MouseIcon       =   "Welcome.frx":28CAC
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   7560
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
      Left            =   3360
      MouseIcon       =   "Welcome.frx":28FB6
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   9000
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
      Left            =   3360
      MouseIcon       =   "Welcome.frx":292C0
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   7200
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
      Left            =   120
      MouseIcon       =   "Welcome.frx":295CA
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   6480
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
      Left            =   120
      MouseIcon       =   "Welcome.frx":298D4
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   7920
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
      Left            =   1440
      MouseIcon       =   "Welcome.frx":29BDE
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   720
      Width           =   5535
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
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   5655
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
      Left            =   120
      MouseIcon       =   "Welcome.frx":2B8A8
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   9000
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
      Left            =   120
      MouseIcon       =   "Welcome.frx":2BBB2
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   6840
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
      Left            =   3360
      MouseIcon       =   "Welcome.frx":2BEBC
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   6120
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
      Left            =   3360
      MouseIcon       =   "Welcome.frx":2C1C6
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   6480
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
      Left            =   3360
      MouseIcon       =   "Welcome.frx":2C4D0
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   8640
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
      Left            =   120
      MouseIcon       =   "Welcome.frx":2C7DA
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   8280
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   7935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Shell "My Calculator 3.exe", vbNormalFocus

End Sub

Private Sub Command10_Click()
Shell "My SL ID Card Reader.exe", vbNormalFocus

End Sub

Private Sub Command11_Click()
Shell "My Text Editor 3.exe", vbNormalFocus

End Sub

Private Sub Command12_Click()
Shell "My AgeInDays.exe", vbNormalFocus

End Sub

Private Sub Command13_Click()
Shell "My Phone Book 3.exe", vbNormalFocus

End Sub

Private Sub Command14_Click()
Shell "My Picture Viewer 2.exe", vbNormalFocus

End Sub

Private Sub Command15_Click()
Shell "My Calendar.exe", vbNormalFocus

End Sub

Private Sub Command16_Click()
Shell "Quiz.exe", vbNormalFocus

End Sub

Private Sub Command17_Click()
Me.Hide
Form2.Show
End Sub

Private Sub Command18_Click()
Shell "My Tamil Keyboard 2.exe", vbNormalFocus

End Sub

Private Sub Command19_Click()
Shell "My Motherboard Quiz.exe", vbNormalFocus

End Sub

Private Sub Command2_Click()
Shell "My Timer 3.exe", vbNormalFocus

End Sub

Private Sub Command20_Click()
Shell "My Unit Converter 2.exe", vbNormalFocus

End Sub

Private Sub Command3_Click()
Shell "My Explorer 6.exe", vbNormalFocus

End Sub

Private Sub Command4_Click()
Shell "My Color Mixer.exe", vbNormalFocus

End Sub

Private Sub Command5_Click()
Shell "My Web Browser.exe", vbNormalFocus

End Sub

Private Sub Command6_Click()
Shell "My Descriptions.exe", vbNormalFocus

End Sub

Private Sub Command7_Click()
Shell "My Chemistry Quiz.exe", vbNormalFocus

End Sub

Private Sub Command8_Click()
Shell "My Media Player 3.exe", vbNormalFocus

End Sub

Private Sub Command9_Click()
Shell "My Media Player 4.exe", vbNormalFocus

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
'MsgBox "If one of my softwares doesn't work properly, please install library files, hyperlink is provided at the bottom"
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
