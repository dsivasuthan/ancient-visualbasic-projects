VERSION 5.00
Begin VB.Form frmSpeedConversion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Speed Conversion"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2850
   Icon            =   "SpeedConversion.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   2850
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt1stUnit 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
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
   Begin VB.TextBox txt2ndUnit 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1320
      Width           =   2655
   End
   Begin VB.ComboBox Combo1stUnit 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "SpeedConversion.frx":57E2
      Left            =   120
      List            =   "SpeedConversion.frx":57EF
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin VB.ComboBox Combo2ndUnit 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "SpeedConversion.frx":580A
      Left            =   120
      List            =   "SpeedConversion.frx":5817
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1680
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Convert"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "="
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "Trebuchet MS"
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
Attribute VB_Name = "frmSpeedConversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Combo1stUnit.Text = "km/hr" Then
If Combo2ndUnit.Text = "km/hr" Then
txt2ndUnit = txt1stUnit * 1
Else
If Combo2ndUnit.Text = "m/sec" Then
txt2ndUnit = (txt1stUnit / 3600) * 1000
Else
If Combo2ndUnit.Text = "mile/hr" Then
txt2ndUnit = (txt1stUnit / 1.609) * 1
End If
End If
End If
End If


If Combo1stUnit.Text = "m/sec" Then
If Combo2ndUnit.Text = "m/sec" Then
txt2ndUnit = txt1stUnit * 1
Else
If Combo2ndUnit.Text = "km/hr" Then
txt2ndUnit = (txt1stUnit / 1000) * 3600
Else
If Combo2ndUnit.Text = "mile/hr" Then
txt2ndUnit = (txt1stUnit / 1609.344) * 3600
End If
End If
End If
End If


If Combo1stUnit.Text = "mile/hr" Then
If Combo2ndUnit.Text = "mile/hr" Then
txt2ndUnit = txt1stUnit * 1
Else
If Combo2ndUnit.Text = "km/hr" Then
txt2ndUnit = (txt1stUnit / 1) * 1.609344
Else
If Combo2ndUnit.Text = "m/sec" Then
txt2ndUnit = (txt1stUnit * 3600) * 1609.344
End If
End If
End If
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
frmWelcomeScreen.Show
End Sub
