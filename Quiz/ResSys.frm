VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Respiratory System Quiz - Dhayalan Sivasuthan"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6255
   Icon            =   "ResSys.frx":0000
   LinkTopic       =   "Form5"
   ScaleHeight     =   5085
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label17 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"ResSys.frx":57E2
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   18
      Top             =   3000
      Width           =   5895
   End
   Begin VB.Line Line1 
      X1              =   4800
      X2              =   5280
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Shape Shape15 
      Height          =   735
      Left            =   4800
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nose"
      DragMode        =   1  'Automatic
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
      Height          =   255
      Left            =   4800
      TabIndex        =   15
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Diaphragm"
      DragMode        =   1  'Automatic
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
      Height          =   255
      Left            =   4800
      TabIndex        =   14
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Lungs"
      DragMode        =   1  'Automatic
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
      Height          =   255
      Left            =   4800
      TabIndex        =   13
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Trachea"
      DragMode        =   1  'Automatic
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
      Height          =   255
      Left            =   4800
      TabIndex        =   12
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Larynx"
      DragMode        =   1  'Automatic
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
      Height          =   255
      Left            =   4800
      TabIndex        =   11
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mouth"
      DragMode        =   1  'Automatic
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
      Height          =   255
      Left            =   4800
      TabIndex        =   10
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bronchi"
      DragMode        =   1  'Automatic
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
      Height          =   255
      Left            =   4800
      TabIndex        =   9
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pharynx"
      DragMode        =   1  'Automatic
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
      Height          =   255
      Left            =   4800
      TabIndex        =   8
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   7
      Left            =   3000
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   6
      Left            =   3000
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   5
      Left            =   3000
      Top             =   840
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   4
      Left            =   120
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   3
      Left            =   3000
      Top             =   480
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   2
      Left            =   360
      Top             =   960
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   1
      Left            =   3240
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   0
      Left            =   3240
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nose"
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
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Diaphragm"
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
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   2250
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Lungs"
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
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Trachea"
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
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Larynx"
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
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mouth"
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
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bronchi"
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
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pharynx"
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
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   945
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   2715
      Left            =   120
      Picture         =   "ResSys.frx":5A76
      Top             =   120
      Width           =   4200
   End
   Begin VB.Label lblTry 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
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
      Left            =   4800
      TabIndex        =   17
      Top             =   555
      Width           =   495
   End
   Begin VB.Label lblMa 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
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
      Left            =   4800
      TabIndex        =   16
      Top             =   195
      Width           =   495
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
Me.Hide
Form4.Show
End Sub



Private Sub Label2_DragDrop(Source As Control, X As Single, Y As Single)
lblTry.Caption = Val(lblTry.Caption) + 1

If Source.Caption = Label2.Caption Then
MsgBox "Correct"

Source.Visible = False
Label2.ForeColor = vbButtonText
lblMa.Caption = Val(lblMa.Caption) + 1
Else
MsgBox "Wrong"
End If

End Sub

Private Sub Label3_DragDrop(Source As Control, X As Single, Y As Single)
lblTry.Caption = Val(lblTry.Caption) + 1

If Source.Caption = Label3.Caption Then
MsgBox "Correct"

Source.Visible = False
Label3.ForeColor = vbButtonText
lblMa.Caption = Val(lblMa.Caption) + 1
Else
MsgBox "Wrong"
End If

End Sub




Private Sub Label4_DragDrop(Source As Control, X As Single, Y As Single)
lblTry.Caption = Val(lblTry.Caption) + 1

If Source.Caption = Label4.Caption Then
MsgBox "Correct"

Source.Visible = False
Label4.ForeColor = vbButtonText
lblMa.Caption = Val(lblMa.Caption) + 1
Else
MsgBox "Wrong"
End If

End Sub

Private Sub Label5_DragDrop(Source As Control, X As Single, Y As Single)
lblTry.Caption = Val(lblTry.Caption) + 1

If Source.Caption = Label5.Caption Then
MsgBox "Correct"

Source.Visible = False
Label5.ForeColor = vbButtonText
lblMa.Caption = Val(lblMa.Caption) + 1
Else
MsgBox "Wrong"
End If

End Sub

Private Sub Label6_DragDrop(Source As Control, X As Single, Y As Single)
lblTry.Caption = Val(lblTry.Caption) + 1

If Source.Caption = Label6.Caption Then
MsgBox "Correct"

Source.Visible = False
Label6.ForeColor = vbButtonText
lblMa.Caption = Val(lblMa.Caption) + 1
Else
MsgBox "Wrong"
End If

End Sub

Private Sub Label7_DragDrop(Source As Control, X As Single, Y As Single)
lblTry.Caption = Val(lblTry.Caption) + 1

If Source.Caption = Label7.Caption Then
MsgBox "Correct"

Source.Visible = False
Label7.ForeColor = vbButtonText
lblMa.Caption = Val(lblMa.Caption) + 1
Else
MsgBox "Wrong"
End If

End Sub

Private Sub Label8_DragDrop(Source As Control, X As Single, Y As Single)
lblTry.Caption = Val(lblTry.Caption) + 1

If Source.Caption = Label8.Caption Then
MsgBox "Correct"

Source.Visible = False
Label8.ForeColor = vbButtonText
lblMa.Caption = Val(lblMa.Caption) + 1
Else
MsgBox "Wrong"
End If

End Sub


Private Sub Label1_DragDrop(Source As Control, X As Single, Y As Single)
lblTry.Caption = Val(lblTry.Caption) + 1

If Source.Caption = Label1.Caption Then
MsgBox "Correct"

Source.Visible = False
Label1.ForeColor = vbButtonText
lblMa.Caption = Val(lblMa.Caption) + 1
Else
MsgBox "Wrong"
End If

End Sub
