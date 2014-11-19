VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form1"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7845
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   7845
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMessage 
      Height          =   285
      Left            =   960
      TabIndex        =   17
      Top             =   4080
      Width           =   4335
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   960
      MaxLength       =   20
      TabIndex        =   16
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Summary"
      Height          =   1215
      Left            =   5400
      TabIndex        =   9
      Top             =   240
      Width           =   3255
      Begin VB.Label lblButtons 
         Caption         =   "Label4"
         Height          =   255
         Left            =   1440
         TabIndex        =   13
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblType 
         Caption         =   "Label3"
         Height          =   255
         Left            =   1440
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Buttons"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Message Type"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Message"
      Height          =   495
      Left            =   2760
      TabIndex        =   8
      Top             =   4560
      Width           =   2535
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Buttons"
      Height          =   3375
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   2535
      Begin VB.OptionButton Option11 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Retry, Cancel"
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Top             =   2880
         Width           =   2175
      End
      Begin VB.OptionButton Option10 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Abort, Retry, Ignore"
         Height          =   315
         Left            =   240
         TabIndex        =   6
         Top             =   2400
         Width           =   2175
      End
      Begin VB.OptionButton Option9 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Yes, No, Cancel"
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   1920
         Width           =   2175
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Yes , No"
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   2175
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "OK , Cancel"
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   2175
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "OK"
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Type of Message"
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      Begin VB.OptionButton optInfo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Informations"
         Height          =   495
         Left            =   840
         TabIndex        =   22
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton optQues 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Question"
         Height          =   495
         Left            =   840
         TabIndex        =   21
         Top             =   960
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optExcla 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Exclamation"
         Height          =   495
         Left            =   840
         TabIndex        =   20
         Top             =   1560
         Width           =   1455
      End
      Begin VB.OptionButton optCriti 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Critical"
         Height          =   495
         Left            =   840
         TabIndex        =   19
         Top             =   2160
         Width           =   1455
      End
      Begin VB.OptionButton optNone 
         BackColor       =   &H00C0C0C0&
         Caption         =   "None"
         Height          =   495
         Left            =   840
         TabIndex        =   18
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Image Image4 
         Height          =   615
         Left            =   180
         Picture         =   "Form1.frx":0000
         Top             =   360
         Width           =   600
      End
      Begin VB.Image Image3 
         Height          =   570
         Left            =   120
         Picture         =   "Form1.frx":0394
         Top             =   960
         Width           =   690
      End
      Begin VB.Image Image2 
         Height          =   705
         Left            =   165
         Picture         =   "Form1.frx":072F
         Top             =   1500
         Width           =   630
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   195
         Picture         =   "Form1.frx":0AC5
         Top             =   2160
         Width           =   585
      End
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Message"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Title"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3720
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Message As String
Dim Buttons As String
Dim TOMB As String

Buttons = lblButtons.Caption
TOMB = lblType.Caption
Message = MsgBox(txtMessage.Text, TOMB + Buttons, txtTitle.Text)

End Sub

Private Sub optCriti_Click()
lblType.Caption = 16
End Sub

Private Sub optExcla_Click()
lblType.Caption = 48
End Sub

Private Sub optInfo_Click()
lblType.Caption = 64
End Sub
Private Sub optNone_Click()
lblType.Caption = ""
End Sub

Private Sub optQues_Click()
lblType.Caption = 32
End Sub

Private Sub Option10_Click()
lblButtons.Caption = vbAbortRetryIgnore
End Sub

Private Sub Option11_Click()
lblButtons.Caption = vbRetryCancel
End Sub

Private Sub Option6_Click()
lblButtons.Caption = vbOKOnly
End Sub

Private Sub Option7_Click()
lblButtons.Caption = vbOKCancel
End Sub

Private Sub Option8_Click()
lblButtons.Caption = vbYesNo
End Sub

Private Sub Option9_Click()
lblButtons.Caption = vbYesNoCancel
End Sub


