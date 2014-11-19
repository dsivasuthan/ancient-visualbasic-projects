VERSION 5.00
Begin VB.Form frmTemperatureConversion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Convert Temperature"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2910
   Icon            =   "TemperatureConversion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   2910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Convert"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
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
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "TemperatureConversion.frx":57E2
      Left            =   120
      List            =   "TemperatureConversion.frx":57EF
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1680
      Width           =   2655
   End
   Begin VB.ComboBox Combo1stUnit 
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
      ItemData        =   "TemperatureConversion.frx":580F
      Left            =   120
      List            =   "TemperatureConversion.frx":581C
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   2655
   End
   Begin VB.TextBox txt2ndUnit 
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
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1320
      Width           =   2655
   End
   Begin VB.TextBox txt1stUnit 
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
         Name            =   "Tahoma"
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
Attribute VB_Name = "frmTemperatureConversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Combo1stUnit.Text = "Celsius" Then
If Combo2ndUnit.Text = "Celsius" Then
 txt2ndUnit = txt1stUnit * 1
Else
If Combo2ndUnit.Text = "Farenheit" Then
  txt2ndUnit = ((9 / 5) * (txt1stUnit)) + 32
Else
If Combo2ndUnit.Text = "kelvin" Then
  txt2ndUnit = txt1stUnit + 273.15
End If
End If
End If
End If



If Combo1stUnit.Text = "Farenheit" Then
If Combo2ndUnit.Text = "Celsius" Then
 txt2ndUnit = (5 / 9) * (txt1stUnit - 32)
Else
If Combo2ndUnit.Text = "Farenheit" Then
  txt2ndUnit = txt1stUnit * 1
Else
If Combo2ndUnit.Text = "kelvin" Then
  txt2ndUnit = "still not found.any comments?"
End If
End If
End If
End If



If Combo1stUnit.Text = "kelvin" Then
If Combo2ndUnit.Text = "Celsius" Then
 txt2ndUnit = txt1stUnit + (-273.15)
Else
If Combo2ndUnit.Text = "Farenheit" Then
  txt2ndUnit = "still not found.any comments?"
Else
If Combo2ndUnit.Text = "kelvin" Then
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
