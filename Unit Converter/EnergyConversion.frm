VERSION 5.00
Begin VB.Form frmEnergyConversion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Convert Energy"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2880
   Icon            =   "EnergyConversion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   2880
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt1stUnit 
      BeginProperty Font 
         Name            =   "Times New Roman"
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
         Name            =   "Times New Roman"
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
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1320
      Width           =   2655
   End
   Begin VB.ComboBox Combo1stUnit 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "EnergyConversion.frx":57E2
      Left            =   120
      List            =   "EnergyConversion.frx":57F2
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   2655
   End
   Begin VB.ComboBox Combo2ndUnit 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "EnergyConversion.frx":580C
      Left            =   120
      List            =   "EnergyConversion.frx":581C
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1680
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Convert"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Times New Roman"
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
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "="
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "Times New Roman"
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
Attribute VB_Name = "frmEnergyConversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Combo1stUnit_Click()
frmLengthConversion.Command1.Caption = "Convert"
End Sub


Private Sub Combo2ndUnit_Click()
frmLengthConversion.Command1.Caption = "Convert"
End Sub

Private Sub Command1_Click()
If Combo1stUnit.Text = "joule" Then
If Combo2ndUnit.Text = "joule" Then
 txt2ndUnit = txt1stUnit * 1
Else
If Combo2ndUnit.Text = "cal" Then
  txt2ndUnit = txt1stUnit * 0.238846
Else
If Combo2ndUnit.Text = "kWh" Then
  txt2ndUnit = txt1stUnit * 2.77777777777778E-07
Else
If Combo2ndUnit.Text = "erg" Then
  txt2ndUnit = txt1stUnit * 10000000
End If
End If
End If
End If
End If



If Combo1stUnit.Text = "cal" Then
If Combo2ndUnit.Text = "joule" Then
 txt2ndUnit = txt1stUnit * 4.1868
Else
If Combo2ndUnit.Text = "cal" Then
  txt2ndUnit = txt1stUnit * 1
Else
If Combo2ndUnit.Text = "kWh" Then
  txt2ndUnit = txt1stUnit * 0.000001163
Else
If Combo2ndUnit.Text = "erg" Then
  txt2ndUnit = txt1stUnit * 41868000
End If
End If
End If
End If
End If



If Combo1stUnit.Text = "kWh" Then
If Combo2ndUnit.Text = "joule" Then
 txt2ndUnit = txt1stUnit * 3600000
Else
If Combo2ndUnit.Text = "cal" Then
  txt2ndUnit = txt1stUnit * 859845.227859
Else
If Combo2ndUnit.Text = "kWh" Then
  txt2ndUnit = txt1stUnit * 1
Else
If Combo2ndUnit.Text = "erg" Then
  txt2ndUnit = txt1stUnit * 36000000000000#
End If
End If
End If
End If
End If



If Combo1stUnit.Text = "erg" Then
If Combo2ndUnit.Text = "joule" Then
 txt2ndUnit = txt1stUnit * 0.0000001
Else
If Combo2ndUnit.Text = "cal" Then
  txt2ndUnit = txt1stUnit * 0.0000000238846
Else
If Combo2ndUnit.Text = "kWh" Then
  txt2ndUnit = txt1stUnit * 1
Else
If Combo2ndUnit.Text = "erg" Then
  txt2ndUnit = txt1stUnit * 2.7777778E-14
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
