VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Heart Quiz - Dhayalan Sivasuthan"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11235
   Icon            =   "Heart.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   11235
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label30 
      Caption         =   $"Heart.frx":57E2
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6495
      Left            =   7800
      TabIndex        =   31
      Top             =   120
      Width           =   3255
   End
   Begin VB.Line Line2 
      Index           =   3
      X1              =   6240
      X2              =   6240
      Y1              =   5640
      Y2              =   6600
   End
   Begin VB.Line Line2 
      Index           =   2
      X1              =   4560
      X2              =   4560
      Y1              =   5640
      Y2              =   6600
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   3120
      X2              =   3120
      Y1              =   5640
      Y2              =   6600
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   1560
      X2              =   1560
      Y1              =   5640
      Y2              =   6600
   End
   Begin VB.Label Label29 
      Caption         =   "Drag and drop each of the label given below to appropriate boxes representing the parts"
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
      Left            =   120
      TabIndex        =   30
      Top             =   120
      Width           =   7815
   End
   Begin VB.Line Line1 
      X1              =   7080
      X2              =   7560
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Shape Shape15 
      Height          =   735
      Left            =   7080
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label20 
      Caption         =   "Septum"
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
      Left            =   6360
      TabIndex        =   19
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label Label19 
      Caption         =   "Myocardium"
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
      Left            =   6360
      TabIndex        =   18
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label Label18 
      Caption         =   "Pulmonary artery"
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
      Left            =   120
      TabIndex        =   17
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label Label17 
      Caption         =   "Aorta"
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
      Left            =   120
      TabIndex        =   16
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label16 
      Caption         =   "Left atrium"
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
      Left            =   120
      TabIndex        =   15
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Label Label15 
      Caption         =   "Aortic valve"
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
      Left            =   3240
      TabIndex        =   14
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label14 
      Caption         =   "Mitral valve"
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
      Left            =   3240
      TabIndex        =   13
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "Left ventricle"
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
      Left            =   3240
      TabIndex        =   12
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Shape Shape14 
      Height          =   255
      Left            =   5760
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Shape Shape13 
      Height          =   255
      Left            =   5760
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Shape Shape12 
      Height          =   255
      Left            =   5760
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Shape Shape11 
      Height          =   255
      Left            =   5760
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Shape Shape10 
      Height          =   255
      Left            =   5760
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Shape Shape9 
      Height          =   255
      Left            =   5760
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Shape Shape8 
      Height          =   255
      Left            =   5760
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Inferior vena cava"
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
      Left            =   1680
      TabIndex        =   5
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Right ventricle"
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
      Left            =   1680
      TabIndex        =   4
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Tricuspid valve"
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
      Left            =   1680
      TabIndex        =   3
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Right atrium"
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
      Left            =   4680
      TabIndex        =   2
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Shape Shape7 
      Height          =   255
      Left            =   120
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Shape Shape6 
      Height          =   255
      Left            =   4680
      Top             =   960
      Width           =   1815
   End
   Begin VB.Shape Shape5 
      Height          =   255
      Left            =   600
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Shape Shape4 
      Height          =   255
      Left            =   600
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Shape Shape3 
      Height          =   255
      Left            =   0
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      Height          =   255
      Left            =   0
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Left            =   0
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Superior vena cava"
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
      Left            =   4680
      TabIndex        =   1
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Pulmonary valve"
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
      Left            =   4680
      TabIndex        =   0
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Superior vena cava"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Inferior vena cava"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Right ventricle"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Tricuspid valve"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Right atrium"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Pulmonary valve"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label28 
      Caption         =   "Septum"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   5880
      TabIndex        =   27
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Label Label27 
      Caption         =   "Myocardium"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   5880
      TabIndex        =   26
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label26 
      Caption         =   "Pulmonary artery"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   5880
      TabIndex        =   25
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label24 
      Caption         =   "Left atrium"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   5880
      TabIndex        =   23
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label23 
      Caption         =   "Aortic valve"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   5880
      TabIndex        =   22
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label22 
      Caption         =   "Mitral valve"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   5880
      TabIndex        =   21
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label21 
      Caption         =   "Left ventricle"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   5880
      TabIndex        =   20
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Aorta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   4800
      TabIndex        =   24
      Top             =   960
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   4935
      Left            =   1800
      Picture         =   "Heart.frx":5D57
      Top             =   600
      Width           =   3990
   End
   Begin VB.Label lblTry 
      Alignment       =   2  'Center
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
      Left            =   7080
      TabIndex        =   29
      Top             =   1035
      Width           =   495
   End
   Begin VB.Label lblMa 
      Alignment       =   2  'Center
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
      Left            =   7080
      TabIndex        =   28
      Top             =   675
      Width           =   495
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
Me.Hide
Form4.Show
End Sub





Private Sub Label10_DragDrop(Source As Control, X As Single, Y As Single)
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



Private Sub Label21_DragDrop(Source As Control, X As Single, Y As Single)

lblTry.Caption = Val(lblTry.Caption) + 1

If Source.Caption = Label21.Caption Then
MsgBox "Correct"

Source.Visible = False
Label21.ForeColor = vbButtonText
lblMa.Caption = Val(lblMa.Caption) + 1
Else
MsgBox "Wrong"
End If
End Sub

Private Sub Label22_DragDrop(Source As Control, X As Single, Y As Single)
If Source.Caption = Label22.Caption Then
MsgBox "Correct"

Source.Visible = False
Label22.ForeColor = vbButtonText
lblMa.Caption = Val(lblMa.Caption) + 1
Else
MsgBox "Wrong"
End If
End Sub

Private Sub Label23_DragDrop(Source As Control, X As Single, Y As Single)

lblTry.Caption = Val(lblTry.Caption) + 1

If Source.Caption = Label23.Caption Then
MsgBox "Correct"

Source.Visible = False
Label23.ForeColor = vbButtonText
lblMa.Caption = Val(lblMa.Caption) + 1
Else
MsgBox "Wrong"
End If
End Sub

Private Sub Label24_DragDrop(Source As Control, X As Single, Y As Single)

lblTry.Caption = Val(lblTry.Caption) + 1

If Source.Caption = Label24.Caption Then
MsgBox "Correct"

Source.Visible = False
Label24.ForeColor = vbButtonText
lblMa.Caption = Val(lblMa.Caption) + 1
Else
MsgBox "Wrong"
End If
End Sub

Private Sub Label25_DragDrop(Source As Control, X As Single, Y As Single)

lblTry.Caption = Val(lblTry.Caption) + 1

If Source.Caption = Label25.Caption Then
MsgBox "Correct"

Source.Visible = False
Label25.ForeColor = vbButtonText
lblMa.Caption = Val(lblMa.Caption) + 1
Else
MsgBox "Wrong"
End If
End Sub

Private Sub Label26_DragDrop(Source As Control, X As Single, Y As Single)

lblTry.Caption = Val(lblTry.Caption) + 1

If Source.Caption = Label26.Caption Then
MsgBox "Correct"

Source.Visible = False
Label26.ForeColor = vbButtonText
lblMa.Caption = Val(lblMa.Caption) + 1
Else
MsgBox "Wrong"
End If
End Sub

Private Sub Label27_DragDrop(Source As Control, X As Single, Y As Single)

lblTry.Caption = Val(lblTry.Caption) + 1

If Source.Caption = Label27.Caption Then
MsgBox "Correct"

Source.Visible = False
Label27.ForeColor = vbButtonText
lblMa.Caption = Val(lblMa.Caption) + 1
Else
MsgBox "Wrong"
End If
End Sub


Private Sub Label28_DragDrop(Source As Control, X As Single, Y As Single)

lblTry.Caption = Val(lblTry.Caption) + 1

If Source.Caption = Label28.Caption Then
MsgBox "Correct"

Source.Visible = False
Label28.ForeColor = vbButtonText

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
