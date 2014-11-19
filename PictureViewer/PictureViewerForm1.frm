VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "My Picture Viewer"
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10005
   Icon            =   "PictureViewerForm1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   10005
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   6240
      Top             =   120
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   2280
      Negotiate       =   -1  'True
      ScaleHeight     =   5
      ScaleLeft       =   1
      ScaleMode       =   0  'User
      ScaleWidth      =   5.02
      TabIndex        =   3
      Top             =   600
      Width           =   7605
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   120
      Pattern         =   "*.jpg;*.bmp;*.gif;*.tif"
      TabIndex        =   2
      Top             =   3000
      Width           =   2055
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2340
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dialog.Show
Me.Enabled = False
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Dir1_Change()

File1.Path = Dir1.Path


End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1.Path = Drive1.Drive

End Sub

Private Sub File1_Click()
If Right(File1.Path, 1) <> "\" Then

filenam = File1.Path + "\" + File1.FileName
Else
filenam = File1.Path + File1.FileName
End If

Picture1.Picture = LoadPicture(filenam)
End Sub

Private Sub Form_Load()
MsgBox "Welcome to Siva's Picture Viewer!", vbOKOnly, "Siva"
End Sub

Private Sub Form_Resize()
On Error Resume Next
Picture1.Width = Me.Width - 2500
Picture1.Height = Me.Height - 1200
File1.Height = Me.Height - 3500
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = True
Dim Ex As String
Ex = MsgBox("Do you really want to quit Siva's Picture Viewer?", vbYesNo, "Siva")
If Ex = vbYes Then
Timer1.Enabled = True
Else
Dialog.Show
Me.Enabled = False
End If
End Sub

Private Sub Timer1_Timer()
If Me.Height > 600 Then
Me.WindowState = 0
Me.Height = Me.Height - 50
Else

End
End If
End Sub
