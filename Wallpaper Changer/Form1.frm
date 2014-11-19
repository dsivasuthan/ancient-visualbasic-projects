VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT3N.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wallpaper Changer - Dhayalan Sivasuthan"
   ClientHeight    =   7245
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   8070
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   8070
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   2880
      TabIndex        =   18
      Top             =   3240
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command7 
      Caption         =   "About"
      Height          =   495
      Left            =   2760
      TabIndex        =   17
      Top             =   6600
      Width           =   2535
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Help"
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   6600
      Width           =   2535
   End
   Begin VB.CommandButton Command5 
      Caption         =   "OK"
      Height          =   495
      Left            =   5400
      TabIndex        =   15
      Top             =   6600
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Apply selected file"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   7200
      Top             =   2880
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Automatically change wallpaper"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   5280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2640
      UseMaskColor    =   -1  'True
      Value           =   2  'Grayed
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Change wallpaper automatically"
      ForeColor       =   &H000080FF&
      Height          =   3015
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   7815
      Begin VB.CommandButton Command4 
         Caption         =   "Remove selected file"
         Height          =   375
         Left            =   3720
         TabIndex        =   14
         Top             =   2520
         Width           =   1935
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   315
         ItemData        =   "Form1.frx":0ECA
         Left            =   1440
         List            =   "Form1.frx":0EF5
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2580
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Apply selected file"
         Height          =   375
         Left            =   5760
         TabIndex        =   10
         Top             =   2520
         Width           =   1935
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   2205
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   7575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "min"
         ForeColor       =   &H0000C0C0&
         Height          =   255
         Left            =   2520
         TabIndex        =   13
         Top             =   2625
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Change wallpaper after every:"
         ForeColor       =   &H0000C0C0&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   2520
         Width           =   1335
      End
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   315
      ItemData        =   "Form1.frx":0F24
      Left            =   2640
      List            =   "Form1.frx":0F37
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   120
      Width           =   2415
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   2040
      Left            =   2640
      Pattern         =   "*.jpg;*.gif;*.bmp;*.wmf"
      TabIndex        =   3
      Top             =   480
      Width           =   2415
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   2565
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2415
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   5280
      Picture         =   "Form1.frx":0F6A
      ScaleHeight     =   2415
      ScaleWidth      =   2700
      TabIndex        =   0
      Top             =   120
      Width           =   2700
      Begin VB.Image Image1 
         Height          =   1695
         Left            =   150
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add selected file to list"
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      Top             =   2640
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Menu About 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Ind As String

Private Sub About_Click()
Command7_Click
End Sub

Private Sub Check1_Click()
If Check1.Value = Unchecked Then
Me.Height = 3855
Command2.Visible = False
Command1.Visible = True
Else
Command2.Visible = True
Command1.Visible = False
Me.Height = 7905
Ind = 0
End If
End Sub

Private Sub Combo1_Change()
If Combo1.Text = "All Picture Files" Then
File1.Pattern = "*.jpg;*.wmf;*.bmp;*.gif"
Else
File1.Pattern = Combo1.Text
End If
End Sub

Private Sub Combo2_Change()
Timer1.Interval = Val(Combo2.Text) * 60
End Sub

Private Sub Command1_Click()
            SavePicture Image1, App.Path & "\wallpaper.bmp"

            SystemParametersInfo 20, 0&, App.Path & "\wallpaper.bmp", &H2 Or &H1
End Sub

Private Sub Command2_Click()
If File1.FileName = "" Then Exit Sub
If Right(Dir1.Path, 1) <> "\" Then
List1.AddItem File1.Path + "\" + File1.FileName
Else
List1.AddItem File1.Path + File1.FileName
End If
End Sub

Private Sub Command3_Click()
            SavePicture Image1, App.Path & "\wallpaper.bmp"

            SystemParametersInfo 20, 0&, App.Path & "\wallpaper.bmp", &H2 Or &H1

End Sub

Private Sub Command4_Click()
On Error Resume Next
List1.RemoveItem List1.ListIndex
End Sub

Private Sub Command5_Click()
Timer1.Enabled = True
Timer1.Interval = Val(Combo2.Text) * 1000 * 60
End Sub

Private Sub Command6_Click()
Form1.Enabled = False
Form2.Show
End Sub

Private Sub Command7_Click()
Me.Enabled = False
Dialog.Show
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
Image1.Picture = LoadPicture(File1.Path + "\" + File1.FileName)
End Sub

Private Sub Form_Load()
Combo1.ListIndex = 0
Combo2.ListIndex = 0
Me.Height = 3855
Ind = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
DEx = True
Dialog.Show
Me.Hide
End Sub

Private Sub List1_Click()
On Error Resume Next
Image1.Picture = LoadPicture(List1.Text)
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
List1.ToolTipText = List1.Text
End Sub

Private Sub Timer1_Timer()

If Check1.Value = Unchecked Then Exit Sub
If List1.ListIndex = List1.ListCount Then
List1.ListIndex = 0
Ind = 0
            SavePicture Image1, App.Path & "\wallpaper.bmp"

            SystemParametersInfo 20, 0&, App.Path & "\wallpaper.bmp", &H2 Or &H1

Else
List1.ListIndex = Ind + 1
            SavePicture Image1, App.Path & "\wallpaper.bmp"

            SystemParametersInfo 20, 0&, App.Path & "\wallpaper.bmp", &H2 Or &H1

End If

End Sub
