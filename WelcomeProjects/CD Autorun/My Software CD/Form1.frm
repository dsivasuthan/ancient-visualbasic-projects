VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12330
   LinkTopic       =   "Form1"
   ScaleHeight     =   10065
   ScaleWidth      =   12330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picmainskin 
      Height          =   9855
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   9795
      ScaleWidth      =   11955
      TabIndex        =   0
      Top             =   0
      Width           =   12015
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "My Software Collection"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   29.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   855
         Left            =   360
         TabIndex        =   2
         Top             =   840
         Width           =   11295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Dhayalan Sivasuthan"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   11295
      End
      Begin VB.Image Image2 
         Height          =   1500
         Left            =   9960
         Picture         =   "Form1.frx":FCE5
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1635
      End
      Begin VB.Image Image1 
         Height          =   1560
         Left            =   360
         Picture         =   "Form1.frx":21A75
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1560
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim WindowRegion As Long
    
    ' I set all these settings here so you won't forget
    ' them and have a non-working demo... Set them in
    ' design time
    picmainskin.ScaleMode = vbPixels
    picmainskin.AutoRedraw = True
    picmainskin.AutoSize = True
    picmainskin.BorderStyle = vbBSNone
    Me.BorderStyle = vbBSNone
        
   ' Set picmainskin.Picture = LoadPicture(App.Path & "\main.bmp")
    
    Me.Width = picmainskin.Width
    Me.Height = picmainskin.Height
    
    WindowRegion = MakeRegion(picmainskin)
    SetWindowRgn Me.hWnd, WindowRegion, True
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



