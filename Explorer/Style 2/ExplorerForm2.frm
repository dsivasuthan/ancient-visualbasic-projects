VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Dialog1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Copy / Move File"
   ClientHeight    =   5895
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6885
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ExplorerForm2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5400
      Picture         =   "ExplorerForm2.frx":058A
      TabIndex        =   12
      ToolTipText     =   "Move the file,shown in source textbox, to the destination shown in target textbox"
      Top             =   5400
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   1680
      TabIndex        =   10
      Top             =   6000
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Move"
      Height          =   375
      Left            =   3960
      Picture         =   "ExplorerForm2.frx":0B14
      TabIndex        =   9
      ToolTipText     =   "Move the file,shown in source textbox, to the destination shown in target textbox"
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Copy"
      Height          =   375
      Left            =   2520
      Picture         =   "ExplorerForm2.frx":109E
      TabIndex        =   8
      ToolTipText     =   "Copy the file,shown in source textbox, to the destination shown in target textbox"
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Create Folder"
      Height          =   375
      Left            =   1080
      Picture         =   "ExplorerForm2.frx":1628
      TabIndex        =   7
      ToolTipText     =   "Make a new folder"
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox txtDesPath 
      Height          =   590
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   4680
      Width           =   5535
   End
   Begin VB.TextBox txtSrcFile 
      Height          =   590
      Left            =   1200
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   4080
      Width           =   5535
   End
   Begin VB.FileListBox File1 
      Enabled         =   0   'False
      Height          =   3015
      Left            =   3480
      TabIndex        =   2
      Top             =   960
      Width           =   3255
   End
   Begin VB.DirListBox Dir1 
      Height          =   2565
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   3255
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Select the destination where the selected file must be copied or moved to and click copy or move button"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   6615
   End
   Begin VB.Label Label2 
      Caption         =   "Destination Path:"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Source File:"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Width           =   1095
   End
End
Attribute VB_Name = "Dialog1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide
Form1.Enabled = True
End Sub

Private Sub Command4_Click()
Dim res As String
res = ShellFileCopy(txtSrcFile.Text, txtDesPath.Text, True)
If res = True Then
MsgBox "File copied"
Else
MsgBox "File not copied"
End If
End Sub

Private Sub Command5_Click()
ShellFileMove txtSrcFile.Text, txtDesPath.Text
MsgBox "file moved"
End Sub

Private Sub Command6_Click()
On Error Resume Next
Dim Fname  As String
Fname = InputBox("Give a name for the new folder", "Folder Name", "New Folder")
MkDir Dir1.Path & "\" & Fname  'for Make new Folder
Dir1.Refresh

End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
txtDesPath.Text = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Unload(Cancel As Integer)
Command1_Click
End Sub
