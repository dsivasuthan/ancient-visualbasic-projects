VERSION 5.00
Begin VB.Form frmPowerConversion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Convert Power"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2850
   Icon            =   "PowerConversion.frx":0000
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
         Name            =   "Tunga"
         Size            =   8.25
         Charset         =   1
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
         Name            =   "Tunga"
         Size            =   8.25
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "PowerConversion.frx":57E2
      Left            =   120
      List            =   "PowerConversion.frx":57EF
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1680
      Width           =   2655
   End
   Begin VB.ComboBox Combo1stUnit 
      BeginProperty Font 
         Name            =   "Tunga"
         Size            =   8.25
         Charset         =   1
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "PowerConversion.frx":580F
      Left            =   120
      List            =   "PowerConversion.frx":581C
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   2655
   End
   Begin VB.TextBox txt2ndUnit 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.000000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Tunga"
         Size            =   8.25
         Charset         =   1
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
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.000000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Tunga"
         Size            =   8.25
         Charset         =   1
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
         Name            =   "Tunga"
         Size            =   24
         Charset         =   1
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
Attribute VB_Name = "frmPowerConversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Command1_Click()
If Combo1stUnit.Text = "watt" Then
If Combo2ndUnit.Text = "watt" Then
txt2ndUnit = txt1stUnit * 1
Else
If Combo2ndUnit.Text = "hp (metric)" Then
txt2ndUnit = txt1stUnit * 0.00136
Else
If Combo2ndUnit.Text = "hp (UK)" Then
txt2ndUnit = txt1stUnit * 0.001341

End If
End If
End If
End If



If Combo1stUnit.Text = "hp (metric)" Then
If Combo2ndUnit.Text = "watt" Then
txt2ndUnit = txt1stUnit * 735.49875
Else
If Combo2ndUnit.Text = "hp (metric)" Then
txt2ndUnit = txt1stUnit * 1
Else
If Combo2ndUnit.Text = "hp (UK)" Then
txt2ndUnit = txt1stUnit * 0.98632

End If
End If
End If
End If


If Combo1stUnit.Text = "hp (UK)" Then
If Combo2ndUnit.Text = "watt" Then
txt2ndUnit = txt1stUnit * 745.699871
Else
If Combo2ndUnit.Text = "hp (metric)" Then
txt2ndUnit = txt1stUnit * 1.01387
Else
If Combo2ndUnit.Text = "hp (UK)" Then
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
