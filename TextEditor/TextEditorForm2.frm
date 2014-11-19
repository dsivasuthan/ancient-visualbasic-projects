VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   1560
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4920
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "TextEditorForm2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Find Again"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Case Sensitive"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Find"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Find what"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Position As String

Private Sub Command1_Click()
Dim compare As Integer

Position = 0
If Check1.Value = 1 Then
    compare = vbBinaryCompare
Else
    compare = vbTextCompare
End If

Position = InStr(Position + 1, Form1.RichTextBox1.Text, Text1.Text, comapre)
If Position > 0 Then
Form1.RichTextBox1.SelStart = Position - 1
Form1.RichTextBox1.SelLength = Len(Text1.Text)
Form1.SetFocus
Else
MsgBox "String not found"
End If

End Sub

Private Sub Command2_Click()
On Error GoTo FindError
Dim compare As Integer


If Check1.Value = 1 Then
    compare = vbBinaryCompare
Else
    compare = vbTextCompare
End If

Position = InStr(Position + 1, Form1.RichTextBox1.Text, Text1.Text, comapre)
If Position > 0 Then
Form1.RichTextBox1.SelStart = Position - 1
Form1.RichTextBox1.SelLength = Len(Text1.Text)
Form1.SetFocus
Else
MsgBox "String not found"
End If


Exit Sub
FindError:
MsgBox "Type a string to find and press Find button"
End Sub

Private Sub Command3_Click()
Unload Me
End Sub
