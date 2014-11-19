VERSION 5.00
Object = "{B4957B60-6071-11CF-A8A0-444553540000}#1.0#0"; "CSIMG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT3N.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "My Picture Viewer 2.1 - Dhayalan Sivasuthan"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12975
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   12975
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "About the Author"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   8160
      Width           =   2655
   End
   Begin ImageViewerCtrl.Viewer Viewer1 
      Height          =   7350
      Left            =   3000
      TabIndex        =   3
      Top             =   240
      Width           =   9720
      _Version        =   65536
      _ExtentX        =   17145
      _ExtentY        =   12965
      _StockProps     =   97
      BorderStyle     =   1
      AutoSize        =   0   'False
      ImageFile       =   ""
      ImageLeft       =   140
      ImageTop        =   100
      ScrollBars      =   3
      Stretch         =   0   'False
      Tiled           =   0   'False
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   495
         Left            =   1800
         TabIndex        =   8
         Top             =   3840
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.ComboBox cmbStretch 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      ItemData        =   "Form1.frx":0CCA
      Left            =   960
      List            =   "Form1.frx":0CD4
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   7200
      Width           =   1815
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   3015
      Left            =   120
      Pattern         =   "*.bmp;*.png;*.gif;*.jpg;*.jpeg;*.wmf"
      TabIndex        =   2
      Top             =   3840
      Width           =   2655
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   3240
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Dhayalan Sivasuthan"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   7800
      Width           =   2655
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FFFF&
      Height          =   22695
      Left            =   2880
      Top             =   120
      Width           =   15
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Stretch:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   150
      TabIndex        =   5
      Top             =   7230
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbStretch_Change()
Viewer1.Stretch = cmbStretch.Text
End Sub

Private Sub cmbStretch_Click()
Viewer1.Stretch = cmbStretch.Text

End Sub


Private Sub Combo1_Change()
Viewer1.TransColor = red
End Sub

Private Sub Command1_Click()
Dialog.Show
Me.Enabled = False
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
Viewer1.ImageFile = File1.Path + "\" + File1.FileName
End Sub

Private Sub Form_Load()
cmbStretch.ListIndex = 1
End Sub

Private Sub Form_Resize()
On Error Resume Next
Viewer1.Height = Me.ScaleHeight - 200
Viewer1.Width = Me.ScaleWidth - 3000
'Shape1.Width = Viewer1.Width + 20
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
