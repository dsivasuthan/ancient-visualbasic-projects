VERSION 5.00
Begin VB.Form frmLengthConversion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Length Conversion"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   Icon            =   "LengthConversion.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   6600
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      ItemData        =   "LengthConversion.frx":57E2
      Left            =   9000
      List            =   "LengthConversion.frx":57E4
      TabIndex        =   8
      Top             =   2640
      Width           =   1215
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      ItemData        =   "LengthConversion.frx":57E6
      Left            =   6600
      List            =   "LengthConversion.frx":57F3
      TabIndex        =   7
      Top             =   2640
      Width           =   1215
   End
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
      TabIndex        =   3
      Top             =   960
      Width           =   6375
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
      ItemData        =   "LengthConversion.frx":580D
      Left            =   3840
      List            =   "LengthConversion.frx":5832
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
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
      ItemData        =   "LengthConversion.frx":589A
      Left            =   120
      List            =   "LengthConversion.frx":58BF
      Style           =   2  'Dropdown List
      TabIndex        =   1
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
      Left            =   3840
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
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
   Begin VB.ComboBox ComboCategory 
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
      ItemData        =   "LengthConversion.frx":5927
      Left            =   6960
      List            =   "LengthConversion.frx":594C
      TabIndex        =   5
      Text            =   "Select the catergory"
      Top             =   960
      Width           =   6375
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
      Left            =   3000
      TabIndex        =   6
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmLengthConversion"
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
Command1.Caption = "Converted"
If Combo1stUnit.Text = "millimetre" Then
If Combo2ndUnit.Text = "millimetre" Then
 txt2ndUnit = txt1stUnit / 1
Else
If Combo2ndUnit.Text = "centimetre" Then
  txt2ndUnit = txt1stUnit / 10
Else
If Combo2ndUnit.Text = "decimetre" Then
  txt2ndUnit = txt1stUnit / 100
Else
If Combo2ndUnit.Text = "metre" Then
    txt2ndUnit = txt1stUnit / 1000
Else
If Combo2ndUnit.Text = "decametre" Then
    txt2ndUnit = txt1stUnit / 10000
Else
If Combo2ndUnit.Text = "hectometre" Then
    txt2ndUnit = txt1stUnit / 100000
Else
If Combo2ndUnit.Text = "kilometre" Then
    txt2ndUnit = txt1stUnit / 1000000
Else
If Combo2ndUnit.Text = "inch" Then
    txt2ndUnit = txt1stUnit * 0.03937
Else
If Combo2ndUnit.Text = "foot" Then
    txt2ndUnit = txt1stUnit * 0.003281
Else
If Combo2ndUnit.Text = "yard" Then
    txt2ndUnit = txt1stUnit * 0.001094
Else
If Combo2ndUnit.Text = "mile" Then
    txt2ndUnit = txt1stUnit * 0.000001

End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If




If Combo1stUnit.Text = "centimetre" Then
If Combo2ndUnit.Text = "millimetre" Then
 txt2ndUnit = txt1stUnit * 10
Else
If Combo2ndUnit.Text = "centimetre" Then
  txt2ndUnit = txt1stUnit / 1
Else
If Combo2ndUnit.Text = "decimetre" Then
  txt2ndUnit = txt1stUnit / 10
Else
If Combo2ndUnit.Text = "metre" Then
    txt2ndUnit = txt1stUnit / 100
Else
If Combo2ndUnit.Text = "decametre" Then
    txt2ndUnit = txt1stUnit / 1000
Else
If Combo2ndUnit.Text = "hectometre" Then
    txt2ndUnit = txt1stUnit / 10000
Else
If Combo2ndUnit.Text = "kilometre" Then
    txt2ndUnit = txt1stUnit / 100000
Else
If Combo2ndUnit.Text = "inch" Then
    txt2ndUnit = txt1stUnit * 0.3937
Else
If Combo2ndUnit.Text = "foot" Then
    txt2ndUnit = txt1stUnit * 0.032808
Else
If Combo2ndUnit.Text = "yard" Then
    txt2ndUnit = txt1stUnit * 0.0010936
Else
If Combo2ndUnit.Text = "mile" Then
    txt2ndUnit = txt1stUnit * 0.000006


End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If


If Combo1stUnit.Text = "decimetre" Then
If Combo2ndUnit.Text = "millimetre" Then
 txt2ndUnit = txt1stUnit * 100
Else
If Combo2ndUnit.Text = "centimetre" Then
  txt2ndUnit = txt1stUnit * 10
Else
If Combo2ndUnit.Text = "decimetre" Then
  txt2ndUnit = txt1stUnit / 1
Else
If Combo2ndUnit.Text = "metre" Then
    txt2ndUnit = txt1stUnit / 10
Else
If Combo2ndUnit.Text = "decametre" Then
    txt2ndUnit = txt1stUnit / 100
Else
If Combo2ndUnit.Text = "hectometre" Then
    txt2ndUnit = txt1stUnit / 1000
Else
If Combo2ndUnit.Text = "kilometre" Then
    txt2ndUnit = txt1stUnit / 10000
Else
If Combo2ndUnit.Text = "inch" Then
    txt2ndUnit = txt1stUnit * 3.937008
Else
If Combo2ndUnit.Text = "foot" Then
    txt2ndUnit = txt1stUnit * 0.328084
Else
If Combo2ndUnit.Text = "yard" Then
    txt2ndUnit = txt1stUnit * 0.109361
Else
If Combo2ndUnit.Text = "mile" Then
    txt2ndUnit = txt1stUnit * 0.000062


End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If


If Combo1stUnit.Text = "metre" Then
If Combo2ndUnit.Text = "millimetre" Then
 txt2ndUnit = txt1stUnit * 1000
Else
If Combo2ndUnit.Text = "centimetre" Then
  txt2ndUnit = txt1stUnit * 100
Else
If Combo2ndUnit.Text = "decimetre" Then
  txt2ndUnit = txt1stUnit * 10
Else
If Combo2ndUnit.Text = "metre" Then
    txt2ndUnit = txt1stUnit / 1
Else
If Combo2ndUnit.Text = "decametre" Then
    txt2ndUnit = txt1stUnit / 10
Else
If Combo2ndUnit.Text = "hectometre" Then
    txt2ndUnit = txt1stUnit / 100
Else
If Combo2ndUnit.Text = "kilometre" Then
    txt2ndUnit = txt1stUnit / 1000
Else
If Combo2ndUnit.Text = "inch" Then
    txt2ndUnit = txt1stUnit * 39.370079
Else
If Combo2ndUnit.Text = "foot" Then
    txt2ndUnit = txt1stUnit * 3.28084
Else
If Combo2ndUnit.Text = "yard" Then
    txt2ndUnit = txt1stUnit * 1.093613
Else
If Combo2ndUnit.Text = "mile" Then
    txt2ndUnit = txt1stUnit * 0.000621


End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If


If Combo1stUnit.Text = "decametre" Then
If Combo2ndUnit.Text = "millimetre" Then
 txt2ndUnit = txt1stUnit * 10000
Else
If Combo2ndUnit.Text = "centimetre" Then
  txt2ndUnit = txt1stUnit * 1000
Else
If Combo2ndUnit.Text = "decimetre" Then
  txt2ndUnit = txt1stUnit * 100
Else
If Combo2ndUnit.Text = "metre" Then
    txt2ndUnit = txt1stUnit * 10
Else
If Combo2ndUnit.Text = "decametre" Then
    txt2ndUnit = txt1stUnit / 1
Else
If Combo2ndUnit.Text = "hectometre" Then
    txt2ndUnit = txt1stUnit / 10
Else
If Combo2ndUnit.Text = "kilometre" Then
    txt2ndUnit = txt1stUnit / 100
Else
If Combo2ndUnit.Text = "inch" Then
    txt2ndUnit = txt1stUnit * 393.700787
Else
If Combo2ndUnit.Text = "foot" Then
    txt2ndUnit = txt1stUnit * 32.808399
Else
If Combo2ndUnit.Text = "yard" Then
    txt2ndUnit = txt1stUnit * 10.936133
Else
If Combo2ndUnit.Text = "mile" Then
    txt2ndUnit = txt1stUnit * 0.006214


End If
End If
End If
End If
End If
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

Private Sub Timer1_Timer()

End Sub

Private Sub txt1stUnit_Change()
frmLengthConversion.Command1.Caption = "Convert"
End Sub

Private Sub txt2ndUnit_Change()
frmLengthConversion.Command1.Caption = "Convert"
End Sub
