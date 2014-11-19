VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Digestive System Quiz - Dhayalan Sivasuthan"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7440
   FillColor       =   &H8000000F&
   ForeColor       =   &H8000000F&
   Icon            =   "DigestSys.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   7440
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label28 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"DigestSys.frx":57E2
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   29
      Top             =   4920
      Width           =   5655
   End
   Begin VB.Label Label27 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dhayalan Sivasuthan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   120
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   12
      Left            =   3480
      Top             =   3960
      Width           =   615
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   11
      Left            =   1320
      Top             =   4320
      Width           =   615
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   10
      Left            =   1200
      Top             =   3840
      Width           =   855
   End
   Begin VB.Shape Shape1 
      Height          =   375
      Index           =   9
      Left            =   120
      Top             =   3360
      Width           =   975
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   8
      Left            =   120
      Top             =   2760
      Width           =   975
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   7
      Left            =   120
      Top             =   2400
      Width           =   975
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   6
      Left            =   120
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   5
      Left            =   960
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   4
      Left            =   3960
      Top             =   960
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   3
      Left            =   3360
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   2
      Left            =   3960
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Index           =   1
      Left            =   3960
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   375
      Index           =   0
      Left            =   4200
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Shape Shape15 
      Height          =   735
      Left            =   6120
      Top             =   120
      Width           =   495
   End
   Begin VB.Line Line1 
      X1              =   6120
      X2              =   6600
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label26 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rectum"
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
      Height          =   255
      Left            =   6120
      TabIndex        =   25
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label25 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Anus"
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
      Height          =   255
      Left            =   6120
      TabIndex        =   24
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label24 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Appendix"
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
      Height          =   255
      Left            =   6120
      TabIndex        =   23
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label23 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Large Intestine"
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
      Height          =   255
      Left            =   6120
      TabIndex        =   22
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label22 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Gall Bladder"
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
      Height          =   255
      Left            =   6120
      TabIndex        =   21
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label21 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Liver"
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
      Height          =   255
      Left            =   6120
      TabIndex        =   20
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label20 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Esophagus"
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
      Height          =   255
      Left            =   6120
      TabIndex        =   19
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label19 
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
      Height          =   255
      Left            =   6120
      TabIndex        =   18
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label18 
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
      Height          =   255
      Left            =   6120
      TabIndex        =   17
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Salivary Glands"
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
      Height          =   255
      Left            =   6120
      TabIndex        =   16
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Stomache"
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
      Height          =   255
      Left            =   6120
      TabIndex        =   15
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pancreas"
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
      Height          =   255
      Left            =   6120
      TabIndex        =   14
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label14 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Small Intestine"
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
      Height          =   255
      Left            =   6120
      TabIndex        =   13
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label13 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rectum"
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
      Left            =   3480
      TabIndex        =   12
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Anus"
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
      Left            =   1320
      TabIndex        =   11
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Appendix"
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
      Left            =   1320
      TabIndex        =   10
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Large Intestine"
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
      TabIndex        =   9
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Gall Bladder"
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
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Liver"
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
      TabIndex        =   7
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Esophagus"
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
      Left            =   480
      TabIndex        =   6
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
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
      Left            =   840
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
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
      Left            =   3960
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Salivary Glands"
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
      Left            =   3360
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Stomache"
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
      Left            =   3960
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pancreas"
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
      Left            =   3960
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Small Intestine"
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
      Left            =   4200
      TabIndex        =   0
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   8340
      Left            =   120
      Picture         =   "DigestSys.frx":58B6
      Top             =   120
      Width           =   4995
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
      Left            =   6120
      TabIndex        =   27
      Top             =   195
      Width           =   495
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
      Left            =   6120
      TabIndex        =   26
      Top             =   555
      Width           =   495
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
Me.Hide
Form4.Show
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

Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
lblTry.Caption = Val(lblTry.Caption) + 1

If Source.Caption = Label10.Caption Then
MsgBox "Correct"

Source.Visible = False
Label10.ForeColor = vbButtonText
lblMa.Caption = Val(lblMa.Caption) + 1
Else
MsgBox "Wrong"
End If

End Sub

Private Sub Label11_DragDrop(Source As Control, X As Single, Y As Single)
lblTry.Caption = Val(lblTry.Caption) + 1

If Source.Caption = Label11.Caption Then
MsgBox "Correct"

Source.Visible = False
Label11.ForeColor = vbButtonText
lblMa.Caption = Val(lblMa.Caption) + 1
Else
MsgBox "Wrong"
End If

End Sub

Private Sub Label12_DragDrop(Source As Control, X As Single, Y As Single)
lblTry.Caption = Val(lblTry.Caption) + 1

If Source.Caption = Label12.Caption Then
MsgBox "Correct"

Source.Visible = False
Label12.ForeColor = vbButtonText
lblMa.Caption = Val(lblMa.Caption) + 1
Else
MsgBox "Wrong"
End If

End Sub

Private Sub Label13_DragDrop(Source As Control, X As Single, Y As Single)
lblTry.Caption = Val(lblTry.Caption) + 1

If Source.Caption = Label13.Caption Then
MsgBox "Correct"

Source.Visible = False
Label13.ForeColor = vbButtonText
lblMa.Caption = Val(lblMa.Caption) + 1
Else
MsgBox "Wrong"
End If

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

Private Sub Label9_DragDrop(Source As Control, X As Single, Y As Single)
lblTry.Caption = Val(lblTry.Caption) + 1

If Source.Caption = Label9.Caption Then
MsgBox "Correct"

Source.Visible = False
Label9.ForeColor = vbButtonText
lblMa.Caption = Val(lblMa.Caption) + 1
Else
MsgBox "Wrong"
End If

End Sub
