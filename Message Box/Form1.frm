VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "D.Siva's Message Box Creator"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5400
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
   ScaleHeight     =   6015
   ScaleWidth      =   5400
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "About"
      Height          =   495
      Left            =   2760
      TabIndex        =   17
      Top             =   5400
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   960
      TabIndex        =   16
      Top             =   4680
      Width           =   4335
   End
   Begin VB.TextBox Text3 
      Height          =   855
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   3720
      Width           =   4335
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Text            =   "Text2"
      Top             =   7200
      Width           =   2895
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Buttons"
      Height          =   2415
      Left            =   2760
      TabIndex        =   7
      Top             =   1080
      Width           =   2535
      Begin VB.OptionButton Option8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Retry / Cancel"
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   1680
         Width           =   2055
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Abort / Retry / Ignore"
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   2055
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "OK / Cancel"
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Yes / No"
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Type of Message"
      Height          =   2415
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2535
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Information"
         Height          =   495
         Left            =   840
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Question"
         Height          =   495
         Left            =   840
         TabIndex        =   5
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Exclamation"
         Height          =   495
         Left            =   840
         TabIndex        =   4
         Top             =   1320
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Critical"
         Height          =   495
         Left            =   840
         TabIndex        =   3
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Image Image4 
         Height          =   570
         Left            =   75
         Picture         =   "Form1.frx":1CCA
         Top             =   750
         Width           =   690
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   165
         Picture         =   "Form1.frx":2065
         Top             =   1770
         Width           =   585
      End
      Begin VB.Image Image2 
         Height          =   705
         Left            =   120
         Picture         =   "Form1.frx":2473
         Top             =   1230
         Width           =   630
      End
      Begin VB.Image Image3 
         Height          =   615
         Left            =   120
         Picture         =   "Form1.frx":2809
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Message"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   5400
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   6720
      Width           =   2895
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5400
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Message Box Creator"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   19
      Top             =   360
      Width           =   4215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Dhayalan Sivasuthan's"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   18
      Top             =   120
      Width           =   4215
   End
   Begin VB.Image Image5 
      Height          =   720
      Left            =   120
      Picture         =   "Form1.frx":2B9D
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Title :"
      Height          =   495
      Left            =   240
      TabIndex        =   15
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Prompt :"
      Height          =   495
      Left            =   240
      TabIndex        =   14
      Top             =   3720
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Select Case Text2
Case "YN"
    Select Case Text1.Text
        Case "vbInformation": MsgBox Text3.Text, vbInformation + vbYesNo, Text4.Text
        Case "vbCritical": MsgBox Text3.Text, vbCritical + vbYesNo, Text4.Text
        Case "vbExclamation": MsgBox Text3.Text, vbExclamation + vbYesNo, Text4.Text
        Case "vbQuestion": MsgBox Text3.Text, vbQuestion + vbYesNo, Text4.Text
    End Select
Case "OC"
    Select Case Text1.Text
        Case "vbInformation": MsgBox Text3.Text, vbInformation + vbOKCancel, Text4.Text
        Case "vbCritical": MsgBox Text3.Text, vbCritical + vbOKCancel, Text4.Text
        Case "vbExclamation": MsgBox Text3.Text, vbExclamation + vbOKCancel, Text4.Text
        Case "vbQuestion": MsgBox Text3.Text, vbQuestion + vbOKCancel, Text4.Text
    End Select

Case "ARI"
    Select Case Text1.Text
        Case "vbInformation": MsgBox Text3.Text, vbInformation + vbAbortRetryIgnore, Text4.Text
        Case "vbCritical": MsgBox Text3.Text, vbCritical + vbAbortRetryIgnore, Text4.Text
        Case "vbExclamation": MsgBox Text3.Text, vbExclamation + vbAbortRetryIgnore, Text4.Text
        Case "vbQuestion": MsgBox Text3.Text, vbQuestion + vbAbortRetryIgnore, Text4.Text
    End Select

Case "RC"
    Select Case Text1.Text
        Case "vbInformation": MsgBox Text3.Text, vbInformation + vbRetryCancel, Text4.Text
        Case "vbCritical": MsgBox Text3.Text, vbCritical + vbRetryCancel, Text4.Text
        Case "vbExclamation": MsgBox Text3.Text, vbExclamation + vbRetryCancel, Text4.Text
        Case "vbQuestion": MsgBox Text3.Text, vbQuestion + vbRetryCancel, Text4.Text
    End Select


End Select
End Sub

Private Sub Command2_Click()
Dialog.Show
Me.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Option1_Click()
Text1.Text = "vbInformation"
End Sub

Private Sub Option2_Click()
Text1.Text = "vbQuestion"
End Sub

Private Sub Option3_Click()
Text1.Text = "vbExclamation"
End Sub

Private Sub Option4_Click()
Text1.Text = "vbCritical"
End Sub

Private Sub Option5_Click()
Text2.Text = "YN"
End Sub

Private Sub Option6_Click()
Text2.Text = "OC"
End Sub

Private Sub Option7_Click()
Text2.Text = "ARI"
End Sub

Private Sub Option8_Click()
Text2.Text = "RC"
End Sub
