VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   2580
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   3060
   ControlBox      =   0   'False
   FontTransparent =   0   'False
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":1CCA
   ScaleHeight     =   2580
   ScaleWidth      =   3060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picMainSkin 
      Height          =   2055
      Left            =   0
      Picture         =   "Form1.frx":3100
      ScaleHeight     =   1995
      ScaleWidth      =   2475
      TabIndex        =   1
      Top             =   0
      Width           =   2535
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   0
         Visible         =   0   'False
         X1              =   915
         X2              =   915
         Y1              =   0
         Y2              =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   1
         Visible         =   0   'False
         X1              =   1005
         X2              =   915
         Y1              =   15
         Y2              =   870
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   2
         Visible         =   0   'False
         X1              =   1110
         X2              =   915
         Y1              =   30
         Y2              =   870
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   3
         Visible         =   0   'False
         X1              =   1185
         X2              =   915
         Y1              =   75
         Y2              =   870
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   4
         Visible         =   0   'False
         X1              =   1275
         X2              =   915
         Y1              =   105
         Y2              =   870
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   5
         Visible         =   0   'False
         X1              =   1350
         X2              =   915
         Y1              =   165
         Y2              =   870
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   6
         Visible         =   0   'False
         X1              =   1455
         X2              =   915
         Y1              =   210
         Y2              =   870
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   7
         Visible         =   0   'False
         X1              =   1545
         X2              =   900
         Y1              =   270
         Y2              =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   8
         Visible         =   0   'False
         X1              =   1635
         X2              =   915
         Y1              =   300
         Y2              =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   9
         Visible         =   0   'False
         X1              =   1680
         X2              =   930
         Y1              =   420
         Y2              =   840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   10
         Visible         =   0   'False
         X1              =   1740
         X2              =   915
         Y1              =   480
         Y2              =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   11
         Visible         =   0   'False
         X1              =   1725
         X2              =   915
         Y1              =   585
         Y2              =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   12
         Visible         =   0   'False
         X1              =   1785
         X2              =   915
         Y1              =   645
         Y2              =   870
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   13
         Visible         =   0   'False
         X1              =   1815
         X2              =   960
         Y1              =   735
         Y2              =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   14
         Visible         =   0   'False
         X1              =   1815
         X2              =   915
         Y1              =   780
         Y2              =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   15
         Visible         =   0   'False
         X1              =   1845
         X2              =   915
         Y1              =   840
         Y2              =   870
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   16
         Visible         =   0   'False
         X1              =   1785
         X2              =   915
         Y1              =   915
         Y2              =   885
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   17
         Visible         =   0   'False
         X1              =   1770
         X2              =   930
         Y1              =   990
         Y2              =   870
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   18
         Visible         =   0   'False
         X1              =   1770
         X2              =   915
         Y1              =   1065
         Y2              =   900
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   19
         Visible         =   0   'False
         X1              =   1740
         X2              =   930
         Y1              =   1155
         Y2              =   900
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   20
         Visible         =   0   'False
         X1              =   1725
         X2              =   930
         Y1              =   1305
         Y2              =   870
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   21
         Visible         =   0   'False
         X1              =   1680
         X2              =   915
         Y1              =   1395
         Y2              =   870
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   22
         Visible         =   0   'False
         X1              =   1590
         X2              =   930
         Y1              =   1440
         Y2              =   915
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   23
         Visible         =   0   'False
         X1              =   1530
         X2              =   915
         Y1              =   1515
         Y2              =   930
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   24
         Visible         =   0   'False
         X1              =   1455
         X2              =   900
         Y1              =   1605
         Y2              =   885
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   25
         Visible         =   0   'False
         X1              =   1335
         X2              =   885
         Y1              =   1650
         Y2              =   915
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   26
         Visible         =   0   'False
         X1              =   1275
         X2              =   900
         Y1              =   1740
         Y2              =   915
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   27
         Visible         =   0   'False
         X1              =   1155
         X2              =   900
         Y1              =   1725
         Y2              =   915
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   28
         Visible         =   0   'False
         X1              =   1080
         X2              =   885
         Y1              =   1755
         Y2              =   930
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   29
         Visible         =   0   'False
         X1              =   1005
         X2              =   900
         Y1              =   1770
         Y2              =   885
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   30
         Visible         =   0   'False
         X1              =   915
         X2              =   915
         Y1              =   1785
         Y2              =   915
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   31
         Visible         =   0   'False
         X1              =   915
         X2              =   840
         Y1              =   885
         Y2              =   1770
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   32
         Visible         =   0   'False
         X1              =   765
         X2              =   900
         Y1              =   1755
         Y2              =   885
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   33
         Visible         =   0   'False
         X1              =   675
         X2              =   900
         Y1              =   1740
         Y2              =   885
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   34
         Visible         =   0   'False
         X1              =   555
         X2              =   900
         Y1              =   1695
         Y2              =   885
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   35
         Visible         =   0   'False
         X1              =   465
         X2              =   915
         Y1              =   1650
         Y2              =   870
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   36
         Visible         =   0   'False
         X1              =   375
         X2              =   900
         Y1              =   1620
         Y2              =   900
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   37
         Visible         =   0   'False
         X1              =   315
         X2              =   915
         Y1              =   1545
         Y2              =   870
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   38
         Visible         =   0   'False
         X1              =   900
         X2              =   255
         Y1              =   885
         Y2              =   1485
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   39
         Visible         =   0   'False
         X1              =   855
         X2              =   210
         Y1              =   885
         Y2              =   1410
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   40
         Visible         =   0   'False
         X1              =   165
         X2              =   930
         Y1              =   1320
         Y2              =   870
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   41
         Visible         =   0   'False
         X1              =   120
         X2              =   915
         Y1              =   1245
         Y2              =   870
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   42
         Visible         =   0   'False
         X1              =   900
         X2              =   75
         Y1              =   885
         Y2              =   1170
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   43
         Visible         =   0   'False
         X1              =   900
         X2              =   30
         Y1              =   885
         Y2              =   1065
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   44
         Visible         =   0   'False
         X1              =   900
         X2              =   15
         Y1              =   885
         Y2              =   975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   45
         Visible         =   0   'False
         X1              =   930
         X2              =   0
         Y1              =   885
         Y2              =   870
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   46
         Visible         =   0   'False
         X1              =   915
         X2              =   15
         Y1              =   870
         Y2              =   780
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   47
         Visible         =   0   'False
         X1              =   870
         X2              =   45
         Y1              =   870
         Y2              =   690
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   48
         Visible         =   0   'False
         X1              =   915
         X2              =   90
         Y1              =   885
         Y2              =   600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   49
         Visible         =   0   'False
         X1              =   900
         X2              =   135
         Y1              =   885
         Y2              =   510
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   50
         Visible         =   0   'False
         X1              =   915
         X2              =   165
         Y1              =   870
         Y2              =   435
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   51
         Visible         =   0   'False
         X1              =   180
         X2              =   885
         Y1              =   330
         Y2              =   885
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   52
         Visible         =   0   'False
         X1              =   900
         X2              =   270
         Y1              =   855
         Y2              =   255
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   53
         Visible         =   0   'False
         X1              =   900
         X2              =   330
         Y1              =   900
         Y2              =   210
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   54
         Visible         =   0   'False
         X1              =   885
         X2              =   435
         Y1              =   855
         Y2              =   150
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   55
         Visible         =   0   'False
         X1              =   870
         X2              =   495
         Y1              =   870
         Y2              =   105
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   56
         Visible         =   0   'False
         X1              =   900
         X2              =   570
         Y1              =   900
         Y2              =   75
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   57
         Visible         =   0   'False
         X1              =   900
         X2              =   690
         Y1              =   900
         Y2              =   45
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   58
         Visible         =   0   'False
         X1              =   900
         X2              =   795
         Y1              =   915
         Y2              =   15
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         Index           =   59
         Visible         =   0   'False
         X1              =   885
         X2              =   870
         Y1              =   855
         Y2              =   0
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000FF&
         Height          =   255
         Left            =   780
         Shape           =   3  'Circle
         Top             =   765
         Width           =   255
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   0
         Visible         =   0   'False
         X1              =   915
         X2              =   915
         Y1              =   915
         Y2              =   405
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   1
         Visible         =   0   'False
         X1              =   915
         X2              =   1185
         Y1              =   915
         Y2              =   480
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   2
         Visible         =   0   'False
         X1              =   915
         X2              =   1320
         Y1              =   915
         Y2              =   675
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   3
         Visible         =   0   'False
         X1              =   915
         X2              =   1410
         Y1              =   885
         Y2              =   900
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   4
         Visible         =   0   'False
         X1              =   1320
         X2              =   915
         Y1              =   1110
         Y2              =   870
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   5
         Visible         =   0   'False
         X1              =   1140
         X2              =   915
         Y1              =   1290
         Y2              =   870
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   6
         Visible         =   0   'False
         X1              =   915
         X2              =   915
         Y1              =   1335
         Y2              =   870
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   7
         Visible         =   0   'False
         X1              =   660
         X2              =   915
         Y1              =   1290
         Y2              =   870
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   8
         Visible         =   0   'False
         X1              =   495
         X2              =   915
         Y1              =   1110
         Y2              =   870
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   9
         Visible         =   0   'False
         X1              =   915
         X2              =   435
         Y1              =   915
         Y2              =   900
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   10
         Visible         =   0   'False
         X1              =   495
         X2              =   915
         Y1              =   630
         Y2              =   885
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Index           =   11
         Visible         =   0   'False
         X1              =   915
         X2              =   690
         Y1              =   900
         Y2              =   480
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4920
      Top             =   4920
   End
   Begin VB.Image Image2 
      Height          =   840
      Left            =   3480
      Picture         =   "Form1.frx":5A35
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   840
   End
   Begin VB.Label Label1 
      Height          =   2295
      Left            =   3480
      TabIndex        =   0
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   2520
      Left            =   2520
      Picture         =   "Form1.frx":15A77
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   2520
   End
   Begin VB.Menu MainMenu 
      Caption         =   "MainMenu"
      Visible         =   0   'False
      Begin VB.Menu Close 
         Caption         =   "Close"
      End
      Begin VB.Menu About 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim Visibilty


Private Sub About_Click()
DEx = False
Dialog.Show
Me.Enabled = False
End Sub

Private Sub Close_Click()
DEx = True
Me.Hide
Dialog.Show
End Sub

Private Sub Form_Load()

Visibilty = Minute(Time)
    Dim WindowRegion As Long
    
    ' I set all these settings here so you won't forget
    ' them and have a non-working demo... Set them in
    ' design time
    picMainSkin.ScaleMode = vbPixels
    picMainSkin.AutoRedraw = True
    picMainSkin.AutoSize = True
    picMainSkin.BorderStyle = vbBSNone
    Me.BorderStyle = vbBSNone
        
'    Set picMainSkin.Picture = LoadPicture(App.Path & "\main.bmp")
    
    Me.Width = picMainSkin.Width
    Me.Height = picMainSkin.Height
    
    WindowRegion = MakeRegion(picMainSkin)
    SetWindowRgn Me.hWnd, WindowRegion, True

End Sub

Private Sub Move_Click()
Me.Caption = "Clock"
End Sub

Private Sub picMainSkin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu MainMenu
End If
End Sub

Private Sub Timer1_Timer()
Line2(1).Visible = False
 Line2(2).Visible = False
 Line2(3).Visible = False
Line2(4).Visible = False
 Line2(5).Visible = False
Line2(6).Visible = False
 Line2(7).Visible = False
 Line2(8).Visible = False
Line2(9).Visible = False
 Line2(10).Visible = False
 Line2(11).Visible = False
Line2(0).Visible = False



 Line1(1).Visible = False
 Line1(2).Visible = False
Line1(3).Visible = False
Line1(4).Visible = False
Line1(5).Visible = False
Line1(6).Visible = False
 Line1(7).Visible = False
Line1(8).Visible = False
Line1(9).Visible = False
Line1(10).Visible = False
Line1(11).Visible = False
Line1(12).Visible = False
Line1(13).Visible = False
Line1(14).Visible = False
Line1(15).Visible = False
Line1(16).Visible = False
Line1(17).Visible = False
Line1(18).Visible = False
Line1(19).Visible = False
Line1(20).Visible = False
Line1(21).Visible = False
Line1(22).Visible = False
Line1(23).Visible = False
Line1(24).Visible = False
Line1(25).Visible = False
Line1(26).Visible = False
Line1(27).Visible = False
Line1(28).Visible = False
Line1(29).Visible = False
Line1(30).Visible = False
Line1(31).Visible = False
Line1(32).Visible = False
Line1(33).Visible = False
Line1(34).Visible = False
Line1(35).Visible = False
Line1(36).Visible = False
Line1(37).Visible = False
Line1(38).Visible = False
Line1(39).Visible = False
Line1(40).Visible = False
Line1(41).Visible = False
Line1(42).Visible = False
Line1(43).Visible = False
 Line1(44).Visible = False
 Line1(45).Visible = False
 Line1(46).Visible = False
 Line1(47).Visible = False
 Line1(48).Visible = False
Line1(49).Visible = False
Line1(50).Visible = False
Line1(51).Visible = False
 Line1(52).Visible = False
Line1(53).Visible = False
Line1(54).Visible = False
Line1(55).Visible = False
Line1(56).Visible = False
Line1(57).Visible = False
Line1(58).Visible = False
Line1(59).Visible = False











If Hour(Time) = 1 Then Line2(1).Visible = True
If Hour(Time) = 2 Then Line2(2).Visible = True
If Hour(Time) = 3 Then Line2(3).Visible = True
If Hour(Time) = 4 Then Line2(4).Visible = True
If Hour(Time) = 5 Then Line2(5).Visible = True
If Hour(Time) = 6 Then Line2(6).Visible = True
If Hour(Time) = 7 Then Line2(7).Visible = True
If Hour(Time) = 8 Then Line2(8).Visible = True
If Hour(Time) = 9 Then Line2(9).Visible = True
If Hour(Time) = 10 Then Line2(10).Visible = True
If Hour(Time) = 11 Then Line2(11).Visible = True
If Hour(Time) = 12 Then Line2(0).Visible = True

If Hour(Time) = 13 Then Line2(1).Visible = True
If Hour(Time) = 14 Then Line2(2).Visible = True
If Hour(Time) = 15 Then Line2(3).Visible = True
If Hour(Time) = 16 Then Line2(4).Visible = True
If Hour(Time) = 17 Then Line2(5).Visible = True
If Hour(Time) = 18 Then Line2(6).Visible = True
If Hour(Time) = 19 Then Line2(7).Visible = True
If Hour(Time) = 20 Then Line2(8).Visible = True
If Hour(Time) = 21 Then Line2(9).Visible = True
If Hour(Time) = 22 Then Line2(10).Visible = True
If Hour(Time) = 23 Then Line2(11).Visible = True
If Hour(Time) = 0 Then Line2(0).Visible = True



If Minute(Time) = 1 Then Line1(1).Visible = True
If Minute(Time) = 2 Then Line1(2).Visible = True
If Minute(Time) = 3 Then Line1(3).Visible = True
If Minute(Time) = 4 Then Line1(4).Visible = True
If Minute(Time) = 5 Then Line1(5).Visible = True
If Minute(Time) = 6 Then Line1(6).Visible = True
If Minute(Time) = 7 Then Line1(7).Visible = True
If Minute(Time) = 8 Then Line1(8).Visible = True
If Minute(Time) = 9 Then Line1(9).Visible = True
If Minute(Time) = 10 Then Line1(10).Visible = True
If Minute(Time) = 11 Then Line1(11).Visible = True
If Minute(Time) = 12 Then Line1(12).Visible = True
If Minute(Time) = 13 Then Line1(13).Visible = True
If Minute(Time) = 14 Then Line1(14).Visible = True
If Minute(Time) = 15 Then Line1(15).Visible = True
If Minute(Time) = 16 Then Line1(16).Visible = True
If Minute(Time) = 17 Then Line1(17).Visible = True
If Minute(Time) = 18 Then Line1(18).Visible = True
If Minute(Time) = 19 Then Line1(19).Visible = True
If Minute(Time) = 20 Then Line1(20).Visible = True
If Minute(Time) = 21 Then Line1(21).Visible = True
If Minute(Time) = 22 Then Line1(22).Visible = True
If Minute(Time) = 23 Then Line1(23).Visible = True
If Minute(Time) = 24 Then Line1(24).Visible = True
If Minute(Time) = 25 Then Line1(25).Visible = True
If Minute(Time) = 26 Then Line1(26).Visible = True
If Minute(Time) = 27 Then Line1(27).Visible = True
If Minute(Time) = 28 Then Line1(28).Visible = True
If Minute(Time) = 29 Then Line1(29).Visible = True
If Minute(Time) = 30 Then Line1(30).Visible = True
If Minute(Time) = 31 Then Line1(31).Visible = True
If Minute(Time) = 32 Then Line1(32).Visible = True
If Minute(Time) = 33 Then Line1(33).Visible = True
If Minute(Time) = 34 Then Line1(34).Visible = True
If Minute(Time) = 35 Then Line1(35).Visible = True
If Minute(Time) = 36 Then Line1(36).Visible = True
If Minute(Time) = 37 Then Line1(37).Visible = True
If Minute(Time) = 38 Then Line1(38).Visible = True
If Minute(Time) = 39 Then Line1(39).Visible = True
If Minute(Time) = 40 Then Line1(40).Visible = True
If Minute(Time) = 41 Then Line1(41).Visible = True
If Minute(Time) = 42 Then Line1(42).Visible = True
If Minute(Time) = 43 Then Line1(43).Visible = True
If Minute(Time) = 44 Then Line1(44).Visible = True
If Minute(Time) = 45 Then Line1(45).Visible = True
If Minute(Time) = 46 Then Line1(46).Visible = True
If Minute(Time) = 47 Then Line1(47).Visible = True
If Minute(Time) = 48 Then Line1(48).Visible = True
If Minute(Time) = 49 Then Line1(49).Visible = True
If Minute(Time) = 50 Then Line1(50).Visible = True
If Minute(Time) = 51 Then Line1(51).Visible = True
If Minute(Time) = 52 Then Line1(52).Visible = True
If Minute(Time) = 53 Then Line1(53).Visible = True
If Minute(Time) = 54 Then Line1(54).Visible = True
If Minute(Time) = 55 Then Line1(55).Visible = True
If Minute(Time) = 56 Then Line1(56).Visible = True
If Minute(Time) = 57 Then Line1(57).Visible = True
If Minute(Time) = 58 Then Line1(58).Visible = True
If Minute(Time) = 59 Then Line1(59).Visible = True




















'Line1(Minute(Time)).Visible = True
'Line1(Minute(Time) - 1).Visible = False
'If Hour(Time) = 13 Then
'Line2(0).Visible = True
'Line2(Hour(Time) - 1).Visible = False
'Exit Sub
'End If
'Line2(Hour(Time)).Visible = True
'Line2(Hour(Time) - 1).Visible = False

'line1(min(time)-1 = Visibilty + 1
'Line1(Visibilty - 1).Visible = False

End Sub
Private Sub picMainSkin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      
      ' Pass the handling of the mouse down message to
      ' the (non-existing really) form caption, so that
      ' the form itself will be dragged when the picture is dragged.
      '
      ' If you have Win 98, Make sure that the "Show window
      ' contents while dragging" display setting is on for nice results.
      
      ReleaseCapture
      SendMessage Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&

End Sub


