VERSION 5.00
Object = "{0002E550-0000-0000-C000-000000000046}#1.0#0"; "OWC10.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10635
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
   ScaleHeight     =   9330
   ScaleWidth      =   10635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command9 
      Height          =   375
      Left            =   2520
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command8 
      Height          =   375
      Left            =   2160
      Picture         =   "Form1.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      Height          =   375
      Left            =   1800
      Picture         =   "Form1.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      Height          =   375
      Left            =   1320
      Picture         =   "Form1.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Height          =   375
      Left            =   840
      Picture         =   "Form1.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Height          =   375
      Left            =   480
      Picture         =   "Form1.frx":1BB2
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Height          =   375
      Left            =   120
      Picture         =   "Form1.frx":213C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Open"
      Height          =   375
      Left            =   7560
      TabIndex        =   2
      Top             =   8760
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   375
      Left            =   9120
      TabIndex        =   1
      Top             =   8760
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5400
      Top             =   9120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Excel Files|*.xls"
   End
   Begin OWC10.Spreadsheet Spreadsheet1 
      Height          =   7860
      Left            =   120
      OleObjectBlob   =   "Form1.frx":26C6
      TabIndex        =   0
      Top             =   720
      Width           =   10440
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
CommonDialog1.ShowSave
Spreadsheet1.Export CommonDialog1.FileName, ssExportActionNone
End Sub

Private Sub Command3_Click()
Spreadsheet1.ActiveCell.Font.Bold = Not Spreadsheet1.ActiveCell.Font.Bold
End Sub

Private Sub Command4_Click()
Spreadsheet1.ActiveCell.Font.Italic = Not Spreadsheet1.ActiveCell.Font.Italic
End Sub

Private Sub Command5_Click()
If Spreadsheet1.ActiveCell.Font.Underline = True Then
  Spreadsheet1.ActiveCell.Font.Underline = False
Else
Spreadsheet1.ActiveCell.Font.Underline = -4142
End If
End Sub

