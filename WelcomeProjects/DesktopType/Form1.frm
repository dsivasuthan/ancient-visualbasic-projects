VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   ClientHeight    =   15330
   ClientLeft      =   105
   ClientTop       =   405
   ClientWidth     =   19170
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   2  'Custom
   Picture         =   "Form1.frx":57E2
   ScaleHeight     =   15330
   ScaleWidth      =   19170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   2175
      Left            =   3960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   64
      Text            =   "Form1.frx":1B665
      Top             =   9840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5880
      Left            =   6240
      TabIndex        =   62
      Top             =   2280
      Width           =   12375
      ExtentX         =   21828
      ExtentY         =   10372
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   17640
      Top             =   14040
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "My Own Programs and Software Collection"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   8040
      TabIndex        =   74
      Top             =   840
      Width           =   9735
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Dhayalan Sivasuthan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   8040
      TabIndex        =   73
      Top             =   360
      Width           =   5775
   End
   Begin VB.Image Image6 
      Height          =   1320
      Left            =   6360
      Picture         =   "Form1.frx":1B6EA
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1440
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   4680
      MouseIcon       =   "Form1.frx":1D22C
      MousePointer    =   99  'Custom
      TabIndex        =   68
      Top             =   7680
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   3240
      MouseIcon       =   "Form1.frx":1DAF6
      MousePointer    =   99  'Custom
      TabIndex        =   67
      Top             =   7680
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   1800
      MouseIcon       =   "Form1.frx":1E3C0
      MousePointer    =   99  'Custom
      TabIndex        =   66
      Top             =   7680
      Width           =   1335
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  'Fixed Single
      Height          =   1920
      Left            =   3975
      Picture         =   "Form1.frx":1EC8A
      Top             =   9240
      Visible         =   0   'False
      Width           =   2010
   End
   Begin VB.Label Labels 
      BackColor       =   &H00000000&
      Caption         =   "www.dsiva.8m.com"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   14
      Left            =   3960
      TabIndex        =   65
      Top             =   9360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Image imgWebForward2 
      Height          =   390
      Left            =   6765
      MouseIcon       =   "Form1.frx":1F479
      MousePointer    =   99  'Custom
      ToolTipText     =   "Forward"
      Top             =   1815
      Width           =   420
   End
   Begin VB.Image imgWebBack2 
      Height          =   390
      Left            =   6360
      MouseIcon       =   "Form1.frx":1FD43
      MousePointer    =   99  'Custom
      ToolTipText     =   "Back"
      Top             =   1800
      Width           =   420
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Home"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7680
      MouseIcon       =   "Form1.frx":2060D
      MousePointer    =   99  'Custom
      TabIndex        =   63
      Top             =   1830
      Width           =   1455
   End
   Begin VB.Image Image5 
      Height          =   345
      Left            =   7635
      Picture         =   "Form1.frx":20ED7
      Top             =   1800
      Width           =   360
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   9360
      MouseIcon       =   "Form1.frx":213E3
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":21CAD
      ToolTipText     =   "End"
      Top             =   1755
      Width           =   480
   End
   Begin VB.Image imgWeb 
      Height          =   555
      Left            =   6240
      Picture         =   "Form1.frx":22977
      Top             =   1725
      Width           =   3780
   End
   Begin VB.Label lblSLCard 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   360
      MouseIcon       =   "Form1.frx":234BA
      MousePointer    =   99  'Custom
      TabIndex        =   60
      Top             =   7680
      Width           =   1335
   End
   Begin VB.Label lblPicViewer 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   4680
      MouseIcon       =   "Form1.frx":23D84
      MousePointer    =   99  'Custom
      TabIndex        =   59
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label lblDidYouKnow 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   1800
      MouseIcon       =   "Form1.frx":2464E
      MousePointer    =   99  'Custom
      TabIndex        =   58
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label lblTamil 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   3240
      MouseIcon       =   "Form1.frx":24F18
      MousePointer    =   99  'Custom
      TabIndex        =   57
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label lblMaths 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   360
      MouseIcon       =   "Form1.frx":257E2
      MousePointer    =   99  'Custom
      TabIndex        =   56
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label lbltimer 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   4680
      MouseIcon       =   "Form1.frx":260AC
      MousePointer    =   99  'Custom
      TabIndex        =   55
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label lblhumanAnatomy 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   3240
      MouseIcon       =   "Form1.frx":26976
      MousePointer    =   99  'Custom
      TabIndex        =   54
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label lblMultiply 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   1800
      MouseIcon       =   "Form1.frx":27240
      MousePointer    =   99  'Custom
      TabIndex        =   53
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label lblMediaSmall 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   360
      MouseIcon       =   "Form1.frx":27B0A
      MousePointer    =   99  'Custom
      TabIndex        =   52
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label lblExplorer 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   4680
      MouseIcon       =   "Form1.frx":283D4
      MousePointer    =   99  'Custom
      TabIndex        =   51
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label lblCheQuiz 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   3240
      MouseIcon       =   "Form1.frx":28C9E
      MousePointer    =   99  'Custom
      TabIndex        =   50
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label lblCalculator 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   1800
      MouseIcon       =   "Form1.frx":29568
      MousePointer    =   99  'Custom
      TabIndex        =   49
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label lblQuiz 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   4680
      MouseIcon       =   "Form1.frx":29E32
      MousePointer    =   99  'Custom
      TabIndex        =   48
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lblUnitConverter 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   4680
      MouseIcon       =   "Form1.frx":2A6FC
      MousePointer    =   99  'Custom
      TabIndex        =   47
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label lblTextEditor 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   360
      MouseIcon       =   "Form1.frx":2AFC6
      MousePointer    =   99  'Custom
      TabIndex        =   46
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label lblPhoneBook 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   3240
      MouseIcon       =   "Form1.frx":2B890
      MousePointer    =   99  'Custom
      TabIndex        =   45
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label lblMotherBoard 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   1800
      MouseIcon       =   "Form1.frx":2C15A
      MousePointer    =   99  'Custom
      TabIndex        =   44
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label lblMediaPlayerBig 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   360
      MouseIcon       =   "Form1.frx":2CA24
      MousePointer    =   99  'Custom
      TabIndex        =   43
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label lblColorMixer 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   3240
      MouseIcon       =   "Form1.frx":2D2EE
      MousePointer    =   99  'Custom
      TabIndex        =   42
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lblCalendar 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   1800
      MouseIcon       =   "Form1.frx":2DBB8
      MousePointer    =   99  'Custom
      TabIndex        =   41
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lblAgeInDays 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   360
      MouseIcon       =   "Form1.frx":2E482
      MousePointer    =   99  'Custom
      TabIndex        =   40
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      BackColor       =   &H00F5F5F5&
      Caption         =   "Multiplication Table"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   19
      Left            =   15240
      TabIndex        =   37
      Top             =   6840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      BackColor       =   &H00F5F5F5&
      Caption         =   "Picture Viewer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   20
      Left            =   13920
      TabIndex        =   36
      Top             =   5640
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      BackColor       =   &H00F5F5F5&
      Caption         =   "Explorer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   21
      Left            =   13920
      TabIndex        =   35
      Top             =   4440
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      Caption         =   "Chemistry Quiz"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   22
      Left            =   13920
      TabIndex        =   34
      Top             =   3240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      BackColor       =   &H00F5F5F5&
      Caption         =   "Calculator 3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   23
      Left            =   13920
      TabIndex        =   33
      Top             =   2040
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      BackColor       =   &H00F5F5F5&
      Caption         =   "MY APPLICATIONS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   24
      Left            =   13920
      TabIndex        =   32
      Top             =   840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      BackColor       =   &H00F5F5F5&
      Caption         =   "Did You Know"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   26
      Left            =   16440
      TabIndex        =   30
      Top             =   720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      BackColor       =   &H00F5F5F5&
      Caption         =   "Media Player (Big)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   29
      Left            =   16440
      TabIndex        =   27
      Top             =   4440
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      Caption         =   "Media Player (Small)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   30
      Left            =   16440
      TabIndex        =   26
      Top             =   3240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      BackColor       =   &H00F5F5F5&
      Caption         =   "Maths Wonder"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   31
      Left            =   16440
      TabIndex        =   25
      Top             =   2040
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      BackColor       =   &H00F5F5F5&
      Caption         =   "MotherBoard Quiz"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   34
      Left            =   16440
      TabIndex        =   22
      Top             =   5640
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   24
      Left            =   17160
      Picture         =   "Form1.frx":2ED4C
      Top             =   3840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   23
      Left            =   17160
      Picture         =   "Form1.frx":2FA16
      Top             =   5040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   21
      Left            =   17160
      Picture         =   "Form1.frx":302E0
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   20
      Left            =   17160
      Picture         =   "Form1.frx":30FAA
      Stretch         =   -1  'True
      Top             =   1440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   16
      Left            =   17160
      Picture         =   "Form1.frx":31534
      Top             =   2640
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   15
      Left            =   14640
      Picture         =   "Form1.frx":321FE
      Top             =   1440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   555
      Index           =   14
      Left            =   14520
      Picture         =   "Form1.frx":32EC8
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   13
      Left            =   14760
      Picture         =   "Form1.frx":33FF8
      Top             =   2640
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   12
      Left            =   14760
      Picture         =   "Form1.frx":34302
      Top             =   3840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   10
      Left            =   14760
      Picture         =   "Form1.frx":34FCC
      Top             =   5040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   16
      Left            =   3915
      TabIndex        =   17
      ToolTipText     =   "Click to know about Me"
      Top             =   14280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Labels 
      BackColor       =   &H00000000&
      Caption         =   " About..."
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   10
      Left            =   4395
      MousePointer    =   99  'Custom
      TabIndex        =   11
      ToolTipText     =   "Click to know about Me"
      Top             =   11265
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Labels 
      BackColor       =   &H00000000&
      Caption         =   "Lael15"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   15
      Left            =   4395
      TabIndex        =   15
      Top             =   11280
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Labels 
      BackColor       =   &H00000000&
      Caption         =   "My Pictures"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   13
      Left            =   4395
      TabIndex        =   14
      Top             =   10320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Labels 
      BackColor       =   &H00000000&
      Caption         =   "My Music"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   12
      Left            =   4395
      TabIndex        =   13
      Top             =   10800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Labels 
      BackColor       =   &H00000000&
      Caption         =   "Label11"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   11
      Left            =   4395
      TabIndex        =   12
      Top             =   9840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   765
      TabIndex        =   16
      Top             =   14910
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgSButton 
      Height          =   540
      Left            =   0
      Picture         =   "Form1.frx":35C96
      ToolTipText     =   "Click Here to begin installing softwares !!"
      Top             =   14760
      Width           =   660
   End
   Begin VB.Label Labels 
      BackColor       =   &H00F5F5F5&
      Caption         =   " Web Browser"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   1080
      TabIndex        =   10
      Top             =   13320
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   7
      Left            =   360
      Picture         =   "Form1.frx":365B7
      Top             =   13200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Labels 
      BackColor       =   &H00F5F5F5&
      Caption         =   " Text Editor"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   1080
      TabIndex        =   9
      Top             =   12720
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Labels 
      BackColor       =   &H00F5F5F5&
      Caption         =   " Multiplication Table"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   1080
      TabIndex        =   8
      Top             =   12120
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Labels 
      BackColor       =   &H00F5F5F5&
      Caption         =   " Picture Viewer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   1080
      TabIndex        =   7
      Top             =   11520
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   6
      Left            =   360
      Picture         =   "Form1.frx":37281
      Top             =   12600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   5
      Left            =   360
      Picture         =   "Form1.frx":37F4B
      Top             =   11400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   4
      Left            =   360
      Picture         =   "Form1.frx":38C15
      Stretch         =   -1  'True
      Top             =   12000
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label Labels 
      BackColor       =   &H00F5F5F5&
      Caption         =   " Explorer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   1080
      TabIndex        =   6
      Top             =   10920
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   3
      Left            =   360
      Picture         =   "Form1.frx":38F67
      Top             =   10800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   2
      Left            =   360
      Picture         =   "Form1.frx":39C31
      Top             =   10200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   435
      Index           =   1
      Left            =   360
      Picture         =   "Form1.frx":39F3B
      Stretch         =   -1  'True
      Top             =   9120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   360
      Picture         =   "Form1.frx":3B06B
      Top             =   9600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Labels 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Chemistry Quiz"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   4
      Left            =   1080
      TabIndex        =   5
      Top             =   10320
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Labels 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   3
      Left            =   4920
      TabIndex        =   3
      ToolTipText     =   "Exit"
      Top             =   14280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Labels 
      BackColor       =   &H00F5F5F5&
      Caption         =   " Calculator 3"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1080
      TabIndex        =   2
      Top             =   9720
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Labels 
      BackColor       =   &H00FFFFFF&
      Caption         =   "   Install Softwares   >"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   13800
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Labels 
      BackColor       =   &H00F5F5F5&
      Caption         =   " MY APPLICATIONS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   9120
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Image imgSButtonBackup 
      Height          =   540
      Left            =   17760
      Picture         =   "Form1.frx":3BD35
      Top             =   13440
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Image imgSButton2 
      Height          =   540
      Left            =   17160
      Picture         =   "Form1.frx":3C656
      Top             =   13440
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Image imgSMenu 
      Height          =   6615
      Left            =   0
      Picture         =   "Form1.frx":3CF77
      Top             =   8280
      Visible         =   0   'False
      Width           =   6105
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Aum"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   16920
      TabIndex        =   4
      Top             =   14940
      Width           =   2175
   End
   Begin VB.Image imgTaskbar 
      Height          =   540
      Left            =   0
      Picture         =   "Form1.frx":44358
      Top             =   14760
      Width           =   19260
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      BackColor       =   &H00F5F5F5&
      Caption         =   "SL ID Card reader"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   37
      Left            =   17760
      TabIndex        =   19
      Top             =   9360
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   27
      Left            =   18480
      Picture         =   "Form1.frx":4621C
      Top             =   8760
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      BackColor       =   &H00F5F5F5&
      Caption         =   "Web Browser"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   17
      Left            =   15240
      TabIndex        =   39
      Top             =   9360
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      BackColor       =   &H00F5F5F5&
      Caption         =   "Text Editor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   18
      Left            =   15240
      TabIndex        =   38
      Top             =   8160
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      BackColor       =   &H00F5F5F5&
      Caption         =   "Age In Days"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   25
      Left            =   15240
      TabIndex        =   31
      Top             =   10440
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      BackColor       =   &H00F5F5F5&
      Caption         =   "Calendar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   27
      Left            =   15240
      TabIndex        =   29
      Top             =   11760
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      BackColor       =   &H00F5F5F5&
      Caption         =   "Color Mixer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   28
      Left            =   15240
      TabIndex        =   28
      Top             =   12960
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      BackColor       =   &H00F5F5F5&
      Caption         =   "General Quiz"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   32
      Left            =   17760
      TabIndex        =   24
      Top             =   8160
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      Caption         =   "PhoneBook"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   33
      Left            =   17760
      TabIndex        =   23
      Top             =   6840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      BackColor       =   &H00F5F5F5&
      Caption         =   "Timer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   35
      Left            =   17760
      TabIndex        =   21
      Top             =   11760
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      Caption         =   "Tamil Keyboard"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   36
      Left            =   17760
      TabIndex        =   20
      Top             =   10440
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Labels 
      Alignment       =   2  'Center
      BackColor       =   &H00F5F5F5&
      Caption         =   "Unit Converter"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   38
      Left            =   17760
      TabIndex        =   18
      Top             =   12960
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   29
      Left            =   18480
      Picture         =   "Form1.frx":46526
      Top             =   12360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   28
      Left            =   18480
      Picture         =   "Form1.frx":46DF0
      Top             =   7560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   26
      Left            =   18480
      Picture         =   "Form1.frx":47ABA
      Top             =   9840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   25
      Left            =   18480
      Picture         =   "Form1.frx":48784
      Top             =   11040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   22
      Left            =   17160
      Picture         =   "Form1.frx":4904E
      Top             =   6240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   19
      Left            =   16080
      Picture         =   "Form1.frx":49D18
      Stretch         =   -1  'True
      Top             =   11040
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   17
      Left            =   16080
      Picture         =   "Form1.frx":4B9E2
      Top             =   12360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   11
      Left            =   14760
      Picture         =   "Form1.frx":4C6AC
      Stretch         =   -1  'True
      Top             =   6240
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   9
      Left            =   16080
      Picture         =   "Form1.frx":4C9FE
      Top             =   7560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   8
      Left            =   16080
      Picture         =   "Form1.frx":4D6C8
      Top             =   8760
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   18
      Left            =   16080
      Picture         =   "Form1.frx":4E392
      Top             =   9840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Click on the start button to install softwares"
      Height          =   495
      Left            =   360
      TabIndex        =   61
      Top             =   14040
      Width           =   2055
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   4680
      MouseIcon       =   "Form1.frx":4F05C
      MousePointer    =   99  'Custom
      TabIndex        =   72
      Top             =   9120
      Width           =   1335
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   3240
      MouseIcon       =   "Form1.frx":4F926
      MousePointer    =   99  'Custom
      TabIndex        =   71
      Top             =   9120
      Width           =   1335
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   1800
      MouseIcon       =   "Form1.frx":501F0
      MousePointer    =   99  'Custom
      TabIndex        =   70
      Top             =   9120
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   360
      MouseIcon       =   "Form1.frx":50ABA
      MousePointer    =   99  'Custom
      TabIndex        =   69
      Top             =   9120
      Width           =   1335
   End
   Begin VB.Image Image4 
      Height          =   10140
      Left            =   240
      Picture         =   "Form1.frx":51384
      Top             =   240
      Width           =   5745
   End
   Begin VB.Menu Apps 
      Caption         =   "Application"
      Begin VB.Menu APIGuide 
         Caption         =   "API Guide"
      End
      Begin VB.Menu Educational 
         Caption         =   "Educational"
      End
      Begin VB.Menu ImageToIcon 
         Caption         =   "Image to Icon"
      End
      Begin VB.Menu LogonStudio 
         Caption         =   "Logon Studio"
      End
      Begin VB.Menu PCSwitchOff 
         Caption         =   "PC Switch Off"
      End
      Begin VB.Menu ScrCapture 
         Caption         =   "Screen Capture"
         Begin VB.Menu CamStu5 
            Caption         =   "Camtasia Studio 5"
         End
         Begin VB.Menu SScrCap 
            Caption         =   "Super Screen Capture"
         End
         Begin VB.Menu ESCV1 
            Caption         =   "Easy Screen Capture"
         End
      End
      Begin VB.Menu PCWizard 
         Caption         =   "PC Wizard"
      End
      Begin VB.Menu AVG 
         Caption         =   "AVG Antivirus Software 8"
      End
      Begin VB.Menu BootSkin 
         Caption         =   "BootSkin"
      End
      Begin VB.Menu PowerPointConveter 
         Caption         =   "PowerPoint Conveter"
      End
      Begin VB.Menu PrismVideoConveter 
         Caption         =   "Prism Video Conveter"
      End
      Begin VB.Menu RecoverMyFiles 
         Caption         =   "Recover My Files"
      End
      Begin VB.Menu Robot 
         Caption         =   "Robot"
      End
      Begin VB.Menu RocketDock 
         Caption         =   "Rocket Dock"
      End
      Begin VB.Menu SkinStudio 
         Caption         =   "Skin Studio"
      End
      Begin VB.Menu SpeakingClock 
         Caption         =   "Speaking Clock"
      End
      Begin VB.Menu SpeedItUp 
         Caption         =   "Speed It Up"
      End
      Begin VB.Menu TweakXP 
         Caption         =   "Tweak XP"
      End
      Begin VB.Menu VBPowerPack 
         Caption         =   "VB Power Pack"
      End
      Begin VB.Menu Verbose 
         Caption         =   "Verbose"
      End
      Begin VB.Menu IconCool 
         Caption         =   "Webshots"
      End
      Begin VB.Menu Webshots 
         Caption         =   "Webshots"
      End
      Begin VB.Menu WindowsMediaPlayer11 
         Caption         =   "Windows Media Player 11"
      End
      Begin VB.Menu WindowsBlinds 
         Caption         =   "Windows Blinds"
      End
      Begin VB.Menu OpenUniverse 
         Caption         =   "Open Universe"
      End
      Begin VB.Menu PlanetApp 
         Caption         =   "Planetary Apprentice"
      End
      Begin VB.Menu recuva 
         Caption         =   "Recuva"
      End
      Begin VB.Menu RevuUnIns 
         Caption         =   "Revo Uninstaller"
      End
      Begin VB.Menu SLScreen 
         Caption         =   "SL Screen saver"
      End
      Begin VB.Menu StartRenamer 
         Caption         =   "Start Renamer"
      End
      Begin VB.Menu StickFigureAni 
         Caption         =   "Stick figure Animator"
      End
      Begin VB.Menu Switch 
         Caption         =   "Switch"
      End
      Begin VB.Menu Periodic 
         Caption         =   "Periodic table"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Const SWP_HIDEWINDOW = &H80
Const SWP_SHOWWINDOW = &H40

Sub Hide_TaskBar() ' Hide The Task Bar
    Dim hwnd1 As Long
    hwnd1 = FindWindow("Shell_traywnd", "")
    Call SetWindowPos(hwnd1, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
End Sub
' ========================================== '
Sub Show_TaskBar()  ' Show The Task Bar
On Error GoTo AnError

    Dim hwnd1 As Long
    hwnd1 = FindWindow("Shell_traywnd", "")
    Call SetWindowPos(hwnd1, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
Exit Sub
AnError:
MsgBox "Ha"
End Sub


Private Sub APIGuide_Click()
On Error GoTo AnError

Shell "Softwares\agsetup.exe", vbNormalFocus
Exit Sub
AnError:
MsgBox "Sorry For the Inconveinience, File was not found", vbInformation

End Sub

Private Sub BootSkin_Click()
On Error GoTo AnError

Shell "Softwares\XP bootskin Maker.exe", vbNormalFocus
Exit Sub
AnError:
MsgBox "Sorry For the Inconveinience, File was not found", vbInformation

End Sub

Private Sub CamStu5_Click()

End Sub

Private Sub Educational_Click()
On Error GoTo AnError

Shell "Softwares\Educational\Edu.exe", vbNormalFocus
Exit Sub
AnError:
MsgBox "Sorry For the Inconveinience, File was not found", vbInformation

End Sub

Private Sub Form_Click()
On Error GoTo AnError
imgSMenu.Visible = False
Labels(0).Visible = False
Labels(1).Visible = False
Labels(2).Visible = False
Labels(3).Visible = False
Labels(4).Visible = False
Labels(5).Visible = False
Labels(6).Visible = False
Labels(7).Visible = False
Labels(8).Visible = False
Labels(9).Visible = False
Labels(10).Visible = False
Labels(11).Visible = False
Labels(12).Visible = False
Labels(13).Visible = False
Labels(14).Visible = False
Labels(15).Visible = False
Labels(16).Visible = False
Text1.Visible = False
Image3.Visible = False

Image1(0).Visible = False
Image1(1).Visible = False
Image1(2).Visible = False
Image1(3).Visible = False
Image1(4).Visible = False
Image1(5).Visible = False
Image1(6).Visible = False
Image1(7).Visible = False
Exit Sub
AnError:
MsgBox "Ha"

End Sub

Private Sub Form_Load()
On Error GoTo AnError

Hide_TaskBar
WebBrowser1.Width = Me.Width - 1800
WebBrowser1.Height = Me.Height - 3000
imgTaskbar.Top = Me.Height - 600
imgSMenu.Top = Me.Height - 7100
imgSButton.Top = Me.Height - 600
Labels(0).Top = Me.Height - 6300
Image1(1).Top = Me.Height - 6300

Labels(1).Top = Me.Height - 1500

Labels(2).Top = Me.Height - 5700
Image1(0).Top = Me.Height - 5700

Labels(3).Top = Me.Height - 1110
Labels(4).Top = Me.Height - 5150
Image1(2).Top = Me.Height - 5150

Labels(5).Top = Me.Height - 4650
Image1(3).Top = Me.Height - 4650

Labels(6).Top = Me.Height - 4050
Image1(5).Top = Me.Height - 4050

Labels(7).Top = Me.Height - 3350
Image1(4).Top = Me.Height - 3350

Labels(8).Top = Me.Height - 2800
Image1(6).Top = Me.Height - 2800

Labels(9).Top = Me.Height - 2300
Image1(7).Top = Me.Height - 2300

Labels(10).Top = Me.Height - 1950
Labels(12).Top = Me.Height - 4600
Labels(13).Top = Me.Height - 5000
Labels(11).Top = Me.Height - 5575
Labels(14).Top = Me.Height - 6000
Labels(15).Top = Me.Height - 4100

Labels(16).Top = Me.Height - 1050
Text1.Top = Me.Height - 6000


Image3.Top = Me.Height - 3675

Label1.Top = Me.Height - 375
Label1.Left = Me.Width - 2200

If Right(App.Path, 1) <> "\" Then
WebBrowser1.Navigate App.Path + "\" + "HTML\HTML.htm"
Else
WebBrowser1.Navigate App.Path + "HTML\HTML.htm"
End If














Exit Sub

AnError:
MsgBox "Ha"

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgSButton.Picture = imgSButtonBackup

End Sub



Private Sub Form_Resize()
imgSMenu.Top = Me.Height - 7100
imgSButton.Top = Me.Height - 600
WebBrowser1.Width = Me.Width - 6800
WebBrowser1.Height = Me.Height - 3000
imgTaskbar.Top = Me.Height - 600
imgSMenu.Top = Me.Height - 7100
imgSButton.Top = Me.Height - 600

Labels(0).Top = Me.Height - 6300
Image1(1).Top = Me.Height - 6300

Labels(1).Top = Me.Height - 1500

Labels(2).Top = Me.Height - 5700
Image1(0).Top = Me.Height - 5700

Labels(3).Top = Me.Height - 1110
Labels(4).Top = Me.Height - 5150
Image1(2).Top = Me.Height - 5150

Labels(5).Top = Me.Height - 4650
Image1(3).Top = Me.Height - 4650

Labels(6).Top = Me.Height - 4050
Image1(5).Top = Me.Height - 4050

Labels(7).Top = Me.Height - 3350
Image1(4).Top = Me.Height - 3350

Labels(8).Top = Me.Height - 2800
Image1(6).Top = Me.Height - 2800

Labels(9).Top = Me.Height - 2300
Image1(7).Top = Me.Height - 2300

Labels(10).Top = Me.Height - 1950
Labels(12).Top = Me.Height - 4600
Labels(13).Top = Me.Height - 5000
Labels(11).Top = Me.Height - 5575
Labels(14).Top = Me.Height - 6000
Labels(15).Top = Me.Height - 4100

Labels(16).Top = Me.Height - 1050
Text1.Top = Me.Height - 6000


Image3.Top = Me.Height - 3675

Label1.Top = Me.Height - 375
Label1.Left = Me.Width - 2200
End Sub

Private Sub Form_Terminate()
On Error GoTo AnError

Show_TaskBar
Exit Sub
AnError:
MsgBox "Ha"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo AnError

Cancel = True
Show_TaskBar
End
Exit Sub
AnError:
MsgBox "Ha"
End Sub



Private Sub IconX_Click()
'Shell "Softwares"
End Sub

Private Sub IconCool_Click()

End Sub

Private Sub Image2_Click()
Show_TaskBar
End
Show_TaskBar
End Sub



Private Sub Image4_Click()
Form_Click
End Sub

Private Sub ImageToIcon_Click()
On Error GoTo AnError

Shell "Softwares\Image to Icon.exe", vbNormalFocus
Exit Sub
AnError:
MsgBox "Sorry For the Inconveinience, File was not found", vbInformation

End Sub

Private Sub imgSButton_Click()
On Error GoTo AnError

If imgSMenu.Visible = True Then
imgSMenu.Visible = False
Labels(0).Visible = False
Labels(1).Visible = False
Labels(2).Visible = False
Labels(3).Visible = False
Labels(4).Visible = False
Labels(5).Visible = False
Labels(6).Visible = False
Labels(7).Visible = False
Labels(8).Visible = False
Labels(9).Visible = False
Labels(10).Visible = False
Labels(11).Visible = False
Labels(12).Visible = False
Labels(13).Visible = False
Labels(14).Visible = False
Labels(15).Visible = False
Labels(16).Visible = False
Text1.Visible = False
Image1(0).Visible = False
Image1(1).Visible = False
Image1(2).Visible = False
Image1(3).Visible = False
Image1(4).Visible = False
Image1(5).Visible = False
Image1(6).Visible = False
Image1(7).Visible = False

Image3.Visible = False

Else
imgSMenu.Visible = True
Labels(0).Visible = True
Labels(1).Visible = True
Labels(2).Visible = True
Labels(3).Visible = True
Labels(4).Visible = True
Labels(5).Visible = True
Labels(6).Visible = True
Labels(7).Visible = True
Labels(8).Visible = True
Labels(9).Visible = True

Image1(0).Visible = True
Image1(1).Visible = True
Image1(2).Visible = True
Image1(3).Visible = True
Image1(4).Visible = True
Image1(5).Visible = True
Image1(6).Visible = True
Image1(7).Visible = True
Labels(10).Visible = True
Labels(11).Visible = True
Labels(12).Visible = True
Labels(13).Visible = True
Labels(14).Visible = True
Labels(15).Visible = True
Labels(16).Visible = True
Text1.Visible = True

Image3.Visible = True
End If

Exit Sub
AnError:
MsgBox "Ha"

End Sub

Private Sub imgSButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgSButton.Picture = imgSButton2.Picture
End Sub







Private Sub imgSMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Labels(0).BackColor = &HF5F5F5
Labels(1).BorderStyle = 0
Labels(2).BackColor = &HF5F5F5
Labels(3).BackColor = &HF5F5F5
Labels(4).BackColor = &HF5F5F5
Labels(5).BackColor = &HF5F5F5
Labels(6).BackColor = &HF5F5F5
Labels(7).BackColor = &HF5F5F5
Labels(8).BackColor = &HF5F5F5
Labels(9).BackColor = &HF5F5F5


End Sub


Private Sub imgTaskbar_Click()
Form_Click
End Sub

Private Sub imgWebBack1_Click()
WebBrowser1.GoBack
End Sub

Private Sub imgTaskbar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgSButton.Picture = imgSButtonBackup
End Sub

Private Sub imgWeb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.FontBold = False
End Sub

Private Sub imgWebBack2_Click()
On Error Resume Next
WebBrowser1.GoBack
End Sub


Private Sub imgWebForward1_Click()
WebBrowser1.GoForward
End Sub

Private Sub imgWebForward2_Click()
On Error Resume Next
WebBrowser1.GoForward
End Sub


Private Sub insPhyTwoPointFive_Click()

End Sub

Private Sub Label10_Click()
On Error GoTo AnError

Shell "My Apps\D.Siva's Unit Converter.exe", vbNormalFocus
Exit Sub
AnError:
MsgBox "Sorry For the Inconveinience, File was not found", vbInformation

End Sub

Private Sub Label11_Click()
On Error GoTo AnError

Shell "My Apps\D.Siva's AgeInDays.exe", vbNormalFocus
Exit Sub
AnError:
MsgBox "Sorry For the Inconveinience, File was not found", vbInformation

End Sub

Private Sub Label4_Click()
WebBrowser1.Navigate App.Path + "\" + "HTML.htm"
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.FontBold = True
End Sub

Private Sub Label5_Click()
On Error GoTo AnError

Shell "My Apps\D.Siva's Media Player 6.2.exe", vbNormalFocus
Exit Sub
AnError:
MsgBox "Sorry For the Inconveinience, File was not found", vbInformation

End Sub

Private Sub Label6_Click()
On Error GoTo AnError

Shell "My Apps\D.Siva's Text Editor 3.exe", vbNormalFocus
Exit Sub
AnError:
MsgBox "Sorry For the Inconveinience, File was not found", vbInformation

End Sub

Private Sub Label7_Click()
On Error GoTo AnError

Shell "My Apps\D.Siva's Tamil  Keyboard 2.1.exe", vbNormalFocus
Exit Sub
AnError:
MsgBox "Sorry For the Inconveinience, File was not found", vbInformation

End Sub

Private Sub Label8_Click()
On Error GoTo AnError

Shell "My Apps\D.Siva's Did You Know 1.1.exe", vbNormalFocus
Exit Sub
AnError:
MsgBox "Sorry For the Inconveinience, File was not found", vbInformation

End Sub

Private Sub Label9_Click()
On Error GoTo AnError

Shell "My Apps\D.Siva's Memory Monitor 1.1.exe", vbNormalFocus
Exit Sub
AnError:
MsgBox "Sorry For the Inconveinience, File was not found", vbInformation

End Sub

Private Sub Labels_Click(Index As Integer)
On Error GoTo AnError

If Index = 1 Then Form1.PopupMenu Apps
If Index = 3 Then
Show_TaskBar
End
End If
If Index = 2 Then lblCalculator_Click
If Index = 4 Then lblCheQuiz_Click
If Index = 5 Then lblExplorer_Click
If Index = 6 Then lblPicViewer_Click
If Index = 7 Then lblMultiply_Click
If Index = 8 Then lblTextEditor_Click
If Index = 9 Then
Shell "Web Browser.exe", vbNormalFocus
End If
If Index = 0 Then
Shell "DSiva'sSoftwares.exe", vbNormalFocus
End If



Exit Sub
AnError:
MsgBox "Ha"
End Sub


Private Sub Labels_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo AnError


If Index = 0 Then
Labels(0).BackColor = &HC0C0C0
Labels(1).BorderStyle = 0
Labels(2).BackColor = &HF5F5F5
Labels(3).BackColor = &HF5F5F5
Labels(4).BackColor = &HF5F5F5
Labels(5).BackColor = &HF5F5F5
Labels(6).BackColor = &HF5F5F5
Labels(7).BackColor = &HF5F5F5
Labels(8).BackColor = &HF5F5F5
Labels(9).BackColor = &HF5F5F5
End If

If Index = 1 Then
Labels(0).BackColor = &HF5F5F5
Labels(1).BorderStyle = 1
Labels(2).BackColor = &HF5F5F5
Labels(3).BackColor = &HF5F5F5
Labels(4).BackColor = &HF5F5F5
Labels(5).BackColor = &HF5F5F5
Labels(6).BackColor = &HF5F5F5
Labels(7).BackColor = &HF5F5F5
Labels(8).BackColor = &HF5F5F5
Labels(9).BackColor = &HF5F5F5

End If

If Index = 2 Then
Labels(0).BackColor = &HF5F5F5
Labels(1).BorderStyle = 0
Labels(2).BackColor = &HC0C0C0
Labels(3).BackColor = &HF5F5F5
Labels(4).BackColor = &HF5F5F5
Labels(5).BackColor = &HF5F5F5
Labels(6).BackColor = &HF5F5F5
Labels(7).BackColor = &HF5F5F5
Labels(8).BackColor = &HF5F5F5
Labels(9).BackColor = &HF5F5F5
End If

If Index = 3 Then
Labels(0).BackColor = &HF5F5F5
Labels(1).BorderStyle = 0
Labels(2).BackColor = &HF5F5F5
Labels(3).BackColor = &HC0C0C0
Labels(4).BackColor = &HF5F5F5
Labels(5).BackColor = &HF5F5F5
Labels(6).BackColor = &HF5F5F5
Labels(7).BackColor = &HF5F5F5
Labels(8).BackColor = &HF5F5F5
Labels(9).BackColor = &HF5F5F5
End If


If Index = 4 Then
Labels(0).BackColor = &HF5F5F5
Labels(1).BorderStyle = 0
Labels(2).BackColor = &HF5F5F5
Labels(3).BackColor = &HF5F5F5
Labels(4).BackColor = &HC0C0C0
Labels(5).BackColor = &HF5F5F5
Labels(6).BackColor = &HF5F5F5
Labels(7).BackColor = &HF5F5F5
Labels(8).BackColor = &HF5F5F5
Labels(9).BackColor = &HF5F5F5
End If


If Index = 5 Then
Labels(0).BackColor = &HF5F5F5
Labels(1).BorderStyle = 0
Labels(2).BackColor = &HF5F5F5
Labels(3).BackColor = &HF5F5F5
Labels(4).BackColor = &HF5F5F5
Labels(5).BackColor = &HC0C0C0
Labels(6).BackColor = &HF5F5F5
Labels(7).BackColor = &HF5F5F5
Labels(8).BackColor = &HF5F5F5
Labels(9).BackColor = &HF5F5F5
End If

If Index = 6 Then
Labels(0).BackColor = &HF5F5F5
Labels(1).BorderStyle = 0
Labels(2).BackColor = &HF5F5F5
Labels(3).BackColor = &HF5F5F5
Labels(4).BackColor = &HF5F5F5
Labels(5).BackColor = &HF5F5F5
Labels(6).BackColor = &HC0C0C0
Labels(7).BackColor = &HF5F5F5
Labels(8).BackColor = &HF5F5F5
Labels(9).BackColor = &HF5F5F5
End If


If Index = 7 Then
Labels(0).BackColor = &HF5F5F5
Labels(1).BorderStyle = 0
Labels(2).BackColor = &HF5F5F5
Labels(3).BackColor = &HF5F5F5
Labels(4).BackColor = &HF5F5F5
Labels(5).BackColor = &HF5F5F5
Labels(6).BackColor = &HF5F5F5
Labels(7).BackColor = &HC0C0C0
Labels(8).BackColor = &HF5F5F5
Labels(9).BackColor = &HF5F5F5
End If


If Index = 8 Then
Labels(0).BackColor = &HF5F5F5
Labels(1).BorderStyle = 0
Labels(2).BackColor = &HF5F5F5
Labels(3).BackColor = &HF5F5F5
Labels(4).BackColor = &HF5F5F5
Labels(5).BackColor = &HF5F5F5
Labels(6).BackColor = &HF5F5F5
Labels(7).BackColor = &HF5F5F5
Labels(8).BackColor = &HC0C0C0
Labels(9).BackColor = &HF5F5F5
End If


If Index = 9 Then
Labels(0).BackColor = &HF5F5F5
Labels(1).BorderStyle = 0
Labels(2).BackColor = &HF5F5F5
Labels(3).BackColor = &HF5F5F5
Labels(4).BackColor = &HF5F5F5
Labels(5).BackColor = &HF5F5F5
Labels(6).BackColor = &HF5F5F5
Labels(7).BackColor = &HF5F5F5
Labels(8).BackColor = &HF5F5F5
Labels(9).BackColor = &HC0C0C0
End If



Exit Sub
AnError:
MsgBox "Ha"

End Sub






















Private Sub lblAgeInDays_Click()
On Error GoTo AnError

Shell "My Apps\D.Siva's Calculator 3.exe", vbNormalFocus
Exit Sub
AnError:
MsgBox "Sorry For the Inconveinience, File was not found", vbInformation
End Sub

Private Sub lblCalculator_Click()
On Error GoTo AnError

Shell "My Apps\D.Siva's Media Player 3.exe", vbNormalFocus
Exit Sub
AnError:
MsgBox "Sorry For the Inconveinience, File was not found", vbInformation


End Sub

Private Sub lblCalendar_Click()
On Error GoTo AnError

Shell "My Apps\Siva's Explorer 7.6.exe", vbNormalFocus
Exit Sub
AnError:
MsgBox "Sorry For the Inconveinience, File was not found", vbInformation

End Sub

Private Sub lblCheQuiz_Click()
On Error GoTo AnError

Shell "My Apps\D.Siva's Multiplication Table.exe", vbNormalFocus
Exit Sub
AnError:
MsgBox "Sorry For the Inconveinience, File was not found", vbInformation


End Sub

Private Sub lblColorMixer_Click()
On Error GoTo AnError

Shell "My Apps\D.Siva's Motherboard Quiz.exe", vbNormalFocus
Exit Sub
AnError:
MsgBox "Sorry For the Inconveinience, File was not found", vbInformation
End Sub

Private Sub lblDidYouKnow_Click()
On Error GoTo AnError

Shell "My Apps\D.Siva's Media Player 5.exe", vbNormalFocus
Exit Sub
AnError:
MsgBox "Sorry For the Inconveinience, File was not found", vbInformation


End Sub

Private Sub lblExplorer_Click()
On Error GoTo AnError

Shell "My Apps\D.Siva's SL ID Card Reader.exe", vbNormalFocus
Exit Sub
AnError:
MsgBox "Sorry For the Inconveinience, File was not found", vbInformation

End Sub

Private Sub lblhumanAnatomy_Click()
On Error GoTo AnError

Shell "My Apps\D.Siva's Picture Viewer 1.exe", vbNormalFocus
Exit Sub
AnError:
MsgBox "Sorry For the Inconveinience, File was not found", vbInformation


End Sub

Private Sub lblMaths_Click()
On Error GoTo AnError

Shell "My Apps\D.Siva's Clock 1.1.exe", vbNormalFocus
Exit Sub
AnError:
MsgBox "Sorry For the Inconveinience, File was not found", vbInformation

End Sub

Private Sub lblMediaPlayerBig_Click()
On Error GoTo AnError

Shell "My Apps\D.Siva's Calendar.exe", vbNormalFocus
Exit Sub
AnError:
MsgBox "Sorry For the Inconveinience, File was not found", vbInformation


End Sub

Private Sub lblMediaSmall_Click()
On Error GoTo AnError

Shell "My Apps\D.Siva's Chemistry Quiz.exe", vbNormalFocus
Exit Sub
AnError:
MsgBox "Sorry For the Inconveinience, File was not found", vbInformation

End Sub

Private Sub lblMotherBoard_Click()
On Error GoTo AnError

Shell "My Apps\D.Siva's Icon Extractor.exe", vbNormalFocus
Exit Sub
AnError:
MsgBox "Sorry For the Inconveinience, File was not found", vbInformation

End Sub

Private Sub lblMultiply_Click()
On Error GoTo AnError

Shell "My Apps\D.Siva's Media Player 4.exe", vbNormalFocus
Exit Sub
AnError:
MsgBox "Sorry For the Inconveinience, File was not found", vbInformation

End Sub

Private Sub lblPhoneBook_Click()
On Error GoTo AnError

Shell "My Apps\Siva's Message Creator.exe", vbNormalFocus
Exit Sub
AnError:
MsgBox "Sorry For the Inconveinience, File was not found", vbInformation


End Sub

Private Sub lblPicViewer_Click()
On Error GoTo AnError

Shell "My Apps\D.Siva's Web Browser.exe", vbNormalFocus
Exit Sub
AnError:
MsgBox "Sorry For the Inconveinience, File was not found", vbInformation


End Sub

Private Sub lblQuiz_Click()
On Error GoTo AnError

Shell "My Apps\D.Siva's Picture Viewer 2.1.exe", vbNormalFocus
Exit Sub
AnError:
MsgBox "Sorry For the Inconveinience, File was not found", vbInformation


End Sub

Private Sub lblSLCard_Click()
On Error GoTo AnError

Shell "My Apps\D.Siva's Color Mixer.exe", vbNormalFocus
Exit Sub
AnError:
MsgBox "Sorry For the Inconveinience, File was not found", vbInformation

End Sub

Private Sub lblTamil_Click()
On Error GoTo AnError

Shell "My Apps\D.Siva's Wallpaper Changer 1.1.exe", vbNormalFocus
Exit Sub
AnError:
MsgBox "Sorry For the Inconveinience, File was not found", vbInformation


End Sub

Private Sub lblTextEditor_Click()
On Error GoTo AnError

Shell "My Apps\D.Siva's Capture 1.1.exe", vbNormalFocus
Exit Sub
AnError:
MsgBox "Sorry For the Inconveinience, File was not found", vbInformation


End Sub

Private Sub lbltimer_Click()
On Error GoTo AnError

Shell "My Apps\D.Siva's Timer 3.1.exe", vbNormalFocus
Exit Sub
AnError:
MsgBox "Sorry For the Inconveinience, File was not found", vbInformation


End Sub

Private Sub lblUnitConverter_Click()
On Error GoTo AnError

Shell "My Apps\D.Siva's Quiz 2.exe", vbNormalFocus
Exit Sub
AnError:
MsgBox "Sorry For the Inconveinience, File was not found", vbInformation


End Sub

Private Sub OLE2_Updated(Code As Integer)

End Sub

Private Sub LogonStudio_Click()
On Error GoTo AnError

Shell "Softwares\LogonStudio_public.exe", vbNormalFocus
Exit Sub
AnError:
MsgBox "Sorry For the Inconveinience, File was not found", vbInformation

End Sub

Private Sub PCSwitchOff_Click()
On Error GoTo AnError

Shell "Softwares\PC Switch off.exe", vbNormalFocus
Exit Sub
AnError:
MsgBox "Sorry For the Inconveinience, File was not found", vbInformation

End Sub

Private Sub PCWizard_Click()

End Sub

Private Sub Periodic_Click()

End Sub

Private Sub recuva_Click()

End Sub

Private Sub SkinStudio_Click()

End Sub

Private Sub StartRenamer_Click()

End Sub

Private Sub Timer1_Timer()
Label1 = Time
End Sub

Private Sub Webshots_Click()

End Sub
