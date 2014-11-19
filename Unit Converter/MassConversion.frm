VERSION 5.00
Begin VB.Form frmMassConversion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mass Conversion"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2880
   Icon            =   "MassConversion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   2880
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
      ItemData        =   "MassConversion.frx":57E2
      Left            =   120
      List            =   "MassConversion.frx":57F2
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
      ItemData        =   "MassConversion.frx":5811
      Left            =   120
      List            =   "MassConversion.frx":5821
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
Attribute VB_Name = "frmMassConversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub Command1_Click()
If Combo1stUnit.Text = "gram" Then
If Combo2ndUnit.Text = "gram" Then
 txt2ndUnit = txt1stUnit * 1
Else
If Combo2ndUnit.Text = "kilogram" Then
  txt2ndUnit = txt1stUnit / 1000
Else
If Combo2ndUnit.Text = "ounce" Then
  txt2ndUnit = txt1stUnit * 0.035273962
Else
If Combo2ndUnit.Text = "lb" Then
    txt2ndUnit = txt1stUnit * 0.002204623

End If
End If
End If
End If
End If



If Combo1stUnit.Text = "kilogram" Then
If Combo2ndUnit.Text = "gram" Then
 txt2ndUnit = txt1stUnit * 1000
Else
If Combo2ndUnit.Text = "kilogram" Then
  txt2ndUnit = txt1stUnit * 1
Else
If Combo2ndUnit.Text = "ounce" Then
  txt2ndUnit = txt1stUnit * 35.273962
Else
If Combo2ndUnit.Text = "lb" Then
    txt2ndUnit = txt1stUnit * 2.204623

End If
End If
End If
End If
End If



If Combo1stUnit.Text = "ounce" Then
If Combo2ndUnit.Text = "gram" Then
 txt2ndUnit = txt1stUnit * 28.349523
Else
If Combo2ndUnit.Text = "kilogram" Then
  txt2ndUnit = txt1stUnit * 0.02835
Else
If Combo2ndUnit.Text = "ounce" Then
  txt2ndUnit = txt1stUnit * 1
Else
If Combo2ndUnit.Text = "lb" Then
    txt2ndUnit = txt1stUnit * 0.0625

End If
End If
End If
End If
End If



If Combo1stUnit.Text = "lb" Then
If Combo2ndUnit.Text = "gram" Then
 txt2ndUnit = txt1stUnit * 453.592374
Else
If Combo2ndUnit.Text = "kilogram" Then
  txt2ndUnit = txt1stUnit * 0.453592
Else
If Combo2ndUnit.Text = "ounce" Then
  txt2ndUnit = txt1stUnit * 16
Else
If Combo2ndUnit.Text = "lb" Then
    txt2ndUnit = txt1stUnit * 1

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
