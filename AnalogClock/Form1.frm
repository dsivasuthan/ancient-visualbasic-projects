VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   6060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   4680
      TabIndex        =   0
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Interval        =   10000
      Left            =   3240
      Top             =   2760
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2760
      Top             =   2760
   End
   Begin VB.Shape Shape3 
      Height          =   1215
      Left            =   1080
      Shape           =   3  'Circle
      Top             =   600
      Width           =   2055
   End
   Begin VB.Shape Shape2 
      Height          =   1455
      Left            =   3960
      Shape           =   3  'Circle
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      Height          =   1455
      Index           =   0
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   960
      Y1              =   3120
      Y2              =   4440
   End
   Begin VB.Line Line2 
      X1              =   840
      X2              =   1680
      Y1              =   2760
      Y2              =   4200
   End
   Begin VB.Line Line1 
      X1              =   1200
      X2              =   2040
      Y1              =   2160
      Y2              =   3840
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim aans As Integer
Dim aanh
Dim hh
Dim aanm As Double
Dim i As Integer
Dim ii As Integer
Dim TWI


Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Command1.Caption = "Quit"
Line1.BorderColor = vbGreen
Line2.BorderColor = vbGreen
Line3.BorderColor = vbGreen
Line1.BorderWidth = 1
Line2.BorderWidth = 4
Line3.BorderWidth = 2
Line1.X1 = 2760
Line1.X2 = 4920
Line2.X1 = 2760
Line2.X2 = 2760
Line3.X1 = 2760
Line3.X2 = 2760
Line1.Y1 = 2520
Line1.Y2 = 2520
Line2.Y1 = 2520
Line2.Y2 = 960
Line3.Y1 = 2520
Line3.Y2 = 4440
Shape3.Width = 255
Shape3.Height = 255
Shape3.Top = 2400
Shape3.Left = 2640
Shape3.BackColor = vbGreen
Shape1(0).Height = 45
Shape1(0).Width = 45
ii = -360
For i = 1 To 59
ii = ii + 6
Load Shape1(i)
Shape1(i).Visible = True
Shape1(i).Left = Line1.X1 + (2160 * (Cos(ii * (22 / 7) / 180)))
Shape1(i).Top = Line1.Y1 + (2160 * (Sin(ii * (22 / 7) / 180)))
Next i
For i = 5 To 57 Step 5
Shape1(i).BorderColor = &HFF0000
Shape1(i).BorderWidth = 2
Next i
aans = -360 + (Second(Now) - 15) * 6
hh = Hour(Now)
If hh = 0 Then hh = 12
If hh > 12 Then
hh = hh - 12
End If
If hh > 3 Then
hh = hh
End If
aanh = -360 + (hh - 3) * 30 + (Minute(Now) * 0.5) + (Second(Now) * 0.0083333333)
If hh = 1 Then
aanh = -60 + (Minute(Now) * 0.5) + (Second(Now) * 0.0083333333)
End If
If hh = 2 Then
aanh = -30 + (Minute(Now) * 0.5) + (Second(Now) * 0.0083333333)
End If
Line2.X2 = Line2.X1 + (1560 * (Cos(aanh * (22 / 7) / 180)))
Line2.Y2 = Line2.Y1 + (1560 * (Sin(aanh * (22 / 7) / 180)))
aanm = -360 + (Minute(Now) - 15) * 6 + (Second(Now) * 0.1)
End Sub

Private Sub Timer1_Timer()
TWI = Second(Now)
If TWI < 15 Then
TWI = 60 - (15 - TWI)
Else
TWI = TWI - 15
End If

aanm = aanm + 0.1
aans = aans + 6
If aans = 6 Then
aans = -354
End If
If aanm = 0.1 Then
aanm = 359.9
End If
Line1.X2 = Line1.X1 + (2160 * (Cos(aans * (22 / 7) / 180)))
Line1.Y2 = Line1.Y1 + (2160 * (Sin(aans * (22 / 7) / 180)))

Line3.X2 = Line3.X1 + (1920 * (Cos(aanm * (22 / 7) / 180)))
Line3.Y2 = Line3.Y1 + (1920 * (Sin(aanm * (22 / 7) / 180)))
Shape1(TWI).BorderColor = &HFFC0C0
If TWI = 0 Then TWI = 60
If Shape1(TWI - 1).BorderWidth = 2 Then
Shape1(TWI - 1).BorderColor = &HFF0000
Exit Sub
End If
Shape1(TWI - 1).BorderColor = &H800000

End Sub

Private Sub Timer2_Timer()

aanh = aanh + 0.083333333
If aanh = 0.083333333 Then
aanh = 359.9166667
End If
Line2.X2 = Line2.X1 + (1560 * (Cos(aanh * (22 / 7) / 180)))
Line2.Y2 = Line2.Y1 + (1560 * (Sin(aanh * (22 / 7) / 180)))
End Sub

