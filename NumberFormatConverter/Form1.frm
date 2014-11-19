VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9045
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   9045
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   3315
      ScaleWidth      =   4875
      TabIndex        =   0
      Top             =   0
      Width           =   4935
      Begin VB.CommandButton Command2 
         Caption         =   "About"
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
         Left            =   240
         TabIndex        =   7
         Top             =   2400
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   1748
         Width           =   2895
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Convert"
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
         Left            =   2400
         TabIndex        =   2
         Top             =   2400
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Image Image2 
         Height          =   540
         Left            =   3960
         Picture         =   "Form1.frx":2D74C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   555
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "Form1.frx":3F4DC
         Top             =   70
         Width           =   480
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Number Format Conveter"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Dhayalan Sivasuthan's"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   45
         Width           =   4455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Roman Numeral"
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
         Left            =   240
         TabIndex        =   5
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Integer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Input a number between 0 to 32675 and press Convert button to turn the given integer into a Roman numeral"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   4095
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************
' Name: ToRoman (Make Integer into a Rom
'     an Numeral!)
' Description:This function will convert
'     a Integer into a Roman Numeral! (Example
'     : 6 --> VI). This function will work
'     from 0 to 32,676 (Integer) and can esial
'     y be made to work for long (upto 9 billi
'     on). This is great! Its quick, and usful
'     ! I have used it before encrypting numbe
'     rs.
' By: Solomon Manalo
'
' Inputs:X - Integer which should be ini
'     talized with the integer you wish to con
'     vert into a roman numeral.
'
' Returns:This function returns a string
'     of the roman numeral converted.
'
' Side Effects:This function only handle
'     s integers upto 32,676 but can be made t
'     o work with long (over 9 billion)
'

Function ToRoman(X As Integer) As String
    ' function provided by Solomon Manalo
    ' code_master_raven@yahoo.com
    ' www.ravensoft.cjb.net
    Dim sFinished As String

    sFinished = String(Int(X / 1000), "M")
    X = X - (Int(X / 1000) * 1000)


    If X >= 900 Then
        sFinished = sFinished & "CM"
    ElseIf X >= 500 And X < 900 Then
        sFinished = sFinished & "D" & String(Int((X - 500) / 100), "C")
    ElseIf X >= 400 And X < 500 Then
        sFinished = sFinished & "CD"
    Else
        sFinished = sFinished & String(Int(X / 100), "C")
    End If
    X = X - (Int(X / 100) * 100)


    If X >= 90 Then
        sFinished = sFinished & "XC"
    ElseIf X >= 50 And X < 90 Then
        sFinished = sFinished & "L" & String(Int((X - 50) / 10), "X")
    ElseIf X >= 40 And X < 50 Then
        sFinished = sFinished & "XL"
    Else
        sFinished = sFinished & String(Int(X / 10), "X")
    End If
    X = X - (Int(X / 10) * 10)


    If X >= 9 Then
        sFinished = sFinished & "IX"
    ElseIf X >= 5 And X < 9 Then
        sFinished = sFinished & "V" & String(Int((X - 5) / 1), "I")
    ElseIf X >= 4 And X < 5 Then
        sFinished = sFinished & "IV"
    Else
        sFinished = sFinished & String(Int(X / 1), "I")
    End If
    ToRoman = sFinished
End Function

Private Sub Command1_Click()
Text2.Text = ToRoman(Text1.Text)
End Sub

Private Sub Command2_Click()
Me.Enabled = False
Dialog.Show
End Sub

Private Sub Form_Load()

End Sub

Private Sub Label2_Click()

End Sub
