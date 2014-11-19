VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Computer Quiz - Dhayalan Sivasuthan"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6120
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
   ScaleHeight     =   1650
   ScaleWidth      =   6120
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Show Form 2"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "<"
      Height          =   495
      Left            =   1440
      TabIndex        =   7
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Next Question >>"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "< OK >"
      Default         =   -1  'True
      Height          =   495
      Left            =   4800
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Top             =   600
      Width           =   4935
   End
   Begin VB.TextBox Text2 
      DataField       =   "Answer"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   2160
      Width           =   4935
   End
   Begin VB.TextBox Text1 
      DataField       =   "Question"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   3855
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\My Documents\VisualBasicProjects\Practice\Quiz\Quiz.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tblQuiz"
      Top             =   3000
      Width           =   1140
   End
   Begin VB.Label Label2 
      Caption         =   "Answer"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "What does this stand for"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text3.Text = Text2.Text Then
MsgBox "Correct"
Data1.Recordset.MoveNext
Text3.Text = ""
Else
MsgBox "Wrong"

End If


Text3.SetFocus
End Sub

Private Sub Command2_Click()
On Error Resume Next
Data1.Recordset.MoveNext
End Sub

Private Sub Command4_Click()
Unload Me
Form2.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Hide
Form4.Show
End Sub
