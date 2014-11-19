VERSION 5.00
Begin VB.Form frmPressureConversion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Convert Pressure"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2850
   Icon            =   "PressureConversion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   2850
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Convert"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   2655
   End
   Begin VB.ComboBox Combo2ndUnit 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "PressureConversion.frx":57E2
      Left            =   120
      List            =   "PressureConversion.frx":57EF
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1680
      Width           =   2655
   End
   Begin VB.ComboBox Combo1stUnit 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "PressureConversion.frx":580D
      Left            =   120
      List            =   "PressureConversion.frx":581A
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   2655
   End
   Begin VB.TextBox txt2ndUnit 
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
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1320
      Width           =   2655
   End
   Begin VB.TextBox txt1stUnit 
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
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "="
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   5
      Top             =   840
      Width           =   615
   End
End
Attribute VB_Name = "frmPressureConversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Combo1stUnit.Text = "pascal" Then
If Combo2ndUnit.Text = "pascal" Then
txt2ndUnit = txt1stUnit * 1
Else
If Combo2ndUnit.Text = "mmHg" Then
txt2ndUnit = txt1stUnit * 0.007501
Else
If Combo2ndUnit.Text = "atmosphere" Then
txt2ndUnit = txt1stUnit * 0.00001

End If
End If
End If
End If



If Combo1stUnit.Text = "atmosphere" Then
If Combo2ndUnit.Text = "pascal" Then
txt2ndUnit = txt1stUnit * 101325
Else
If Combo2ndUnit.Text = "atmosphere" Then
txt2ndUnit = txt1stUnit * 1
Else
If Combo2ndUnit.Text = "mmHg" Then
txt2ndUnit = txt1stUnit * 759.999892
End If
End If
End If
End If



If Combo1stUnit.Text = "mmHg" Then
If Combo2ndUnit.Text = "pascal" Then
txt2ndUnit = txt1stUnit * 133.322387
Else
If Combo2ndUnit.Text = "atmosphere" Then
txt2ndUnit = txt1stUnit * 0.001316
Else
If Combo2ndUnit.Text = "mmHg" Then
txt2ndUnit = txt1stUnit * 1
End If
End If
End If
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
frmWelcomeScreen.Show
End Sub
