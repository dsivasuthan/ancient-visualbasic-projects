VERSION 5.00
Begin VB.Form frmAreaConversion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Convert Area"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   Icon            =   "AreaConversion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   6630
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   4
      Top             =   120
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
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
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
      ItemData        =   "AreaConversion.frx":57E2
      Left            =   120
      List            =   "AreaConversion.frx":57F5
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
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
      ItemData        =   "AreaConversion.frx":5835
      Left            =   3840
      List            =   "AreaConversion.frx":5848
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
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
      TabIndex        =   0
      Top             =   960
      Width           =   6375
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
      Left            =   3000
      TabIndex        =   5
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmAreaConversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Combo1stUnit.Text = "square metre" Then
If Combo2ndUnit.Text = "square metre" Then
 txt2ndUnit = txt1stUnit * 1
Else
If Combo2ndUnit.Text = "hectare" Then
  txt2ndUnit = txt1stUnit * 0.0001
Else
If Combo2ndUnit.Text = "acre" Then
  txt2ndUnit = txt1stUnit * 0.000247
Else
If Combo2ndUnit.Text = "square mile" Then
  txt2ndUnit = txt1stUnit * 0.0000003861022
Else
If Combo2ndUnit.Text = "square kilometre" Then
  txt2ndUnit = txt1stUnit * 0.000001

End If
End If
End If
End If
End If
End If


If Combo1stUnit.Text = "square kilometre" Then
If Combo2ndUnit.Text = "square metre" Then
 txt2ndUnit = txt1stUnit * 1000000
Else
If Combo2ndUnit.Text = "hectare" Then
  txt2ndUnit = txt1stUnit * 100
Else
If Combo2ndUnit.Text = "acre" Then
  txt2ndUnit = txt1stUnit * 247.105381
Else
If Combo2ndUnit.Text = "square mile" Then
  txt2ndUnit = txt1stUnit * 0.386102
Else
If Combo2ndUnit.Text = "square kilometre" Then
  txt2ndUnit = txt1stUnit * 1

End If
End If
End If
End If
End If
End If



If Combo1stUnit.Text = "hectare" Then
If Combo2ndUnit.Text = "square metre" Then
 txt2ndUnit = txt1stUnit * 10000
Else
If Combo2ndUnit.Text = "hectare" Then
  txt2ndUnit = txt1stUnit * 1
Else
If Combo2ndUnit.Text = "acre" Then
  txt2ndUnit = txt1stUnit * 2.471054
Else
If Combo2ndUnit.Text = "square mile" Then
  txt2ndUnit = txt1stUnit * 0.003861
Else
If Combo2ndUnit.Text = "square kilometre" Then
  txt2ndUnit = txt1stUnit * 0.01

End If
End If
End If
End If
End If
End If


If Combo1stUnit.Text = "acre" Then
If Combo2ndUnit.Text = "square metre" Then
 txt2ndUnit = txt1stUnit * 4046.856422
Else
If Combo2ndUnit.Text = "hectare" Then
  txt2ndUnit = txt1stUnit * 0.404686
Else
If Combo2ndUnit.Text = "acre" Then
  txt2ndUnit = txt1stUnit * 1
Else
If Combo2ndUnit.Text = "square mile" Then
  txt2ndUnit = txt1stUnit * 0.001563
Else
If Combo2ndUnit.Text = "square kilometre" Then
  txt2ndUnit = txt1stUnit * 0.004047

End If
End If
End If
End If
End If
End If


If Combo1stUnit.Text = "square mile" Then
If Combo2ndUnit.Text = "square metre" Then
 txt2ndUnit = txt1stUnit * 2589988.110336
Else
If Combo2ndUnit.Text = "hectare" Then
  txt2ndUnit = txt1stUnit * 258.998811
Else
If Combo2ndUnit.Text = "acre" Then
  txt2ndUnit = txt1stUnit * 640
Else
If Combo2ndUnit.Text = "square mile" Then
  txt2ndUnit = txt1stUnit * 1
Else
If Combo2ndUnit.Text = "square kilometre" Then
  txt2ndUnit = txt1stUnit * 2.589988

End If
End If
End If
End If
End If
End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
Unload Me
frmWelcomeScreen.Show

End Sub
