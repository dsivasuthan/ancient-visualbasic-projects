VERSION 5.00
Begin VB.Form frmVolumeConversion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Convert Volume"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2910
   Icon            =   "VolumeConversion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
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
      ItemData        =   "VolumeConversion.frx":57E2
      Left            =   120
      List            =   "VolumeConversion.frx":57F8
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
      ItemData        =   "VolumeConversion.frx":5841
      Left            =   120
      List            =   "VolumeConversion.frx":5857
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
Attribute VB_Name = "frmVolumeConversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Combo1stUnit.Text = "litre (cu dm)" Then
If Combo2ndUnit.Text = "litre (cu dm)" Then
 txt2ndUnit = txt1stUnit * 1
Else
If Combo2ndUnit.Text = "cu metre" Then
  txt2ndUnit = txt1stUnit * 0.001
Else
If Combo2ndUnit.Text = "cu inch" Then
  txt2ndUnit = txt1stUnit * 61.023744
Else
If Combo2ndUnit.Text = "cu foot" Then
  txt2ndUnit = txt1stUnit * 0.035315
Else
If Combo2ndUnit.Text = "gallon (UK)" Then
  txt2ndUnit = txt1stUnit * 0.219969
Else
If Combo2ndUnit.Text = "gallon (US)" Then
  txt2ndUnit = txt1stUnit * 0.264172
End If
End If
End If
End If
End If
End If
End If



If Combo1stUnit.Text = "cu metre" Then
If Combo2ndUnit.Text = "litre (cu dm)" Then
 txt2ndUnit = txt1stUnit * 1000
Else
If Combo2ndUnit.Text = "cu metre" Then
  txt2ndUnit = txt1stUnit * 1
Else
If Combo2ndUnit.Text = "cu inch" Then
  txt2ndUnit = txt1stUnit * 61023.744095
Else
If Combo2ndUnit.Text = "cu foot" Then
  txt2ndUnit = txt1stUnit * 35.314667
Else
If Combo2ndUnit.Text = "gallon (UK)" Then
  txt2ndUnit = txt1stUnit * 219.969248
Else
If Combo2ndUnit.Text = "gallon (US)" Then
  txt2ndUnit = txt1stUnit * 264.172052
End If
End If
End If
End If
End If
End If
End If


If Combo1stUnit.Text = "cu inch" Then
If Combo2ndUnit.Text = "litre (cu dm)" Then
 txt2ndUnit = txt1stUnit * 0.016387
Else
If Combo2ndUnit.Text = "cu metre" Then
  txt2ndUnit = txt1stUnit * 0.000016
Else
If Combo2ndUnit.Text = "cu inch" Then
  txt2ndUnit = txt1stUnit * 1
Else
If Combo2ndUnit.Text = "cu foot" Then
  txt2ndUnit = txt1stUnit * 0.000579
Else
If Combo2ndUnit.Text = "gallon (UK)" Then
  txt2ndUnit = txt1stUnit * 0.003605
Else
If Combo2ndUnit.Text = "gallon (US)" Then
  txt2ndUnit = txt1stUnit * 0.004329
End If
End If
End If
End If
End If
End If
End If


If Combo1stUnit.Text = "cu foot" Then
If Combo2ndUnit.Text = "litre (cu dm)" Then
 txt2ndUnit = txt1stUnit * 28.316847
Else
If Combo2ndUnit.Text = "cu metre" Then
  txt2ndUnit = txt1stUnit * 0.028317
Else
If Combo2ndUnit.Text = "cu inch" Then
  txt2ndUnit = txt1stUnit * 1728
Else
If Combo2ndUnit.Text = "cu foot" Then
  txt2ndUnit = txt1stUnit * 1
Else
If Combo2ndUnit.Text = "gallon (UK)" Then
  txt2ndUnit = txt1stUnit * 6.228835
Else
If Combo2ndUnit.Text = "gallon (US)" Then
  txt2ndUnit = txt1stUnit * 7.480519
End If
End If
End If
End If
End If
End If
End If


If Combo1stUnit.Text = "gallon (UK)" Then
If Combo2ndUnit.Text = "litre (cu dm)" Then
 txt2ndUnit = txt1stUnit * 4.54609
Else
If Combo2ndUnit.Text = "cu metre" Then
  txt2ndUnit = txt1stUnit * 0.004546
Else
If Combo2ndUnit.Text = "cu inch" Then
  txt2ndUnit = txt1stUnit * 277.419433
Else
If Combo2ndUnit.Text = "cu foot" Then
  txt2ndUnit = txt1stUnit * 0.160544
Else
If Combo2ndUnit.Text = "gallon (UK)" Then
  txt2ndUnit = txt1stUnit * 1
Else
If Combo2ndUnit.Text = "gallon (US)" Then
  txt2ndUnit = txt1stUnit * 1.20095
End If
End If
End If
End If
End If
End If
End If


If Combo1stUnit.Text = "gallon (US)" Then
If Combo2ndUnit.Text = "litre (cu dm)" Then
 txt2ndUnit = txt1stUnit * 3.785412
Else
If Combo2ndUnit.Text = "cu metre" Then
  txt2ndUnit = txt1stUnit * 0.003785
Else
If Combo2ndUnit.Text = "cu inch" Then
  txt2ndUnit = txt1stUnit * 231#
Else
If Combo2ndUnit.Text = "cu foot" Then
  txt2ndUnit = txt1stUnit * 0.133681
Else
If Combo2ndUnit.Text = "gallon (UK)" Then
  txt2ndUnit = txt1stUnit * 0.832674
Else
If Combo2ndUnit.Text = "gallon (US)" Then
  txt2ndUnit = txt1stUnit * 1
End If
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
