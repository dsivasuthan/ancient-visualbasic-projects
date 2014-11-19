VERSION 5.00
Begin VB.Form frmDensityConversion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Convert Density"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2880
   Icon            =   "DensityConversion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   2880
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
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1320
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
      ItemData        =   "DensityConversion.frx":57E2
      Left            =   120
      List            =   "DensityConversion.frx":57EF
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
      ItemData        =   "DensityConversion.frx":5818
      Left            =   120
      List            =   "DensityConversion.frx":5825
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
      Top             =   2160
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
Attribute VB_Name = "frmDensityConversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Combo1stUnit_Click()
frmLengthConversion.Command1.Caption = "Convert"
End Sub

Private Sub Command1_Click()
If Combo1stUnit.Text = "kg/cu metre" Then
If Combo2ndUnit.Text = "kg/cu metre" Then
 txt2ndUnit = txt1stUnit * 1
Else
If Combo2ndUnit.Text = "gram/cu cm" Then
  txt2ndUnit = txt1stUnit * 0.001
Else
If Combo2ndUnit.Text = "lb/cu inch" Then
  txt2ndUnit = txt1stUnit * 0.000036
End If
End If
End If
End If



If Combo1stUnit.Text = "gram/cu cm" Then
If Combo2ndUnit.Text = "kg/cu metre" Then
 txt2ndUnit = txt1stUnit * 1000
Else
If Combo2ndUnit.Text = "gram/cu cm" Then
  txt2ndUnit = txt1stUnit * 1
Else
If Combo2ndUnit.Text = "lb/cu inch" Then
  txt2ndUnit = txt1stUnit * 0.036127
End If
End If
End If
End If


If Combo1stUnit.Text = "lb/cu inch" Then
If Combo2ndUnit.Text = "kg/cu metre" Then
 txt2ndUnit = txt1stUnit * 27679.90498
Else
If Combo2ndUnit.Text = "gram/cu cm" Then
  txt2ndUnit = txt1stUnit * 27.679905
Else
If Combo2ndUnit.Text = "lb/cu inch" Then
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
