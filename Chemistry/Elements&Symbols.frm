VERSION 5.00
Begin VB.Form frmSymbols 
   Caption         =   "Common Elements & Symbols"
   ClientHeight    =   7320
   ClientLeft      =   8970
   ClientTop       =   8475
   ClientWidth     =   9630
   Icon            =   "Elements&Symbols.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Elements&Symbols.frx":57E2
   ScaleHeight     =   7320
   ScaleWidth      =   9630
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "End"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   7920
      Picture         =   "Elements&Symbols.frx":C757
      Style           =   1  'Graphical
      TabIndex        =   145
      TabStop         =   0   'False
      Top             =   6600
      Width           =   1575
   End
   Begin VB.TextBox Zn2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9060
      Locked          =   -1  'True
      TabIndex        =   144
      TabStop         =   0   'False
      Top             =   6000
      Width           =   465
   End
   Begin VB.TextBox Y2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9060
      Locked          =   -1  'True
      TabIndex        =   143
      TabStop         =   0   'False
      Top             =   5520
      Width           =   465
   End
   Begin VB.TextBox Xe2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9060
      Locked          =   -1  'True
      TabIndex        =   142
      TabStop         =   0   'False
      Top             =   5040
      Width           =   465
   End
   Begin VB.TextBox Va2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9060
      Locked          =   -1  'True
      TabIndex        =   141
      TabStop         =   0   'False
      Top             =   4560
      Width           =   465
   End
   Begin VB.TextBox U2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9060
      Locked          =   -1  'True
      TabIndex        =   140
      TabStop         =   0   'False
      Top             =   4080
      Width           =   465
   End
   Begin VB.TextBox Te2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9060
      Locked          =   -1  'True
      TabIndex        =   139
      TabStop         =   0   'False
      Top             =   3600
      Width           =   465
   End
   Begin VB.TextBox Ti2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9060
      Locked          =   -1  'True
      TabIndex        =   138
      TabStop         =   0   'False
      Top             =   3120
      Width           =   465
   End
   Begin VB.TextBox Th2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9060
      Locked          =   -1  'True
      TabIndex        =   137
      TabStop         =   0   'False
      Top             =   2640
      Width           =   465
   End
   Begin VB.TextBox Sn2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9060
      Locked          =   -1  'True
      TabIndex        =   136
      TabStop         =   0   'False
      Top             =   2160
      Width           =   465
   End
   Begin VB.TextBox S2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9060
      Locked          =   -1  'True
      TabIndex        =   135
      TabStop         =   0   'False
      Top             =   1680
      Width           =   465
   End
   Begin VB.TextBox Na2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9060
      Locked          =   -1  'True
      TabIndex        =   134
      TabStop         =   0   'False
      Top             =   1200
      Width           =   465
   End
   Begin VB.TextBox Sr2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9060
      Locked          =   -1  'True
      TabIndex        =   133
      TabStop         =   0   'False
      Top             =   720
      Width           =   465
   End
   Begin VB.TextBox Ag2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6675
      Locked          =   -1  'True
      TabIndex        =   132
      TabStop         =   0   'False
      Top             =   6000
      Width           =   465
   End
   Begin VB.TextBox Rn2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6675
      Locked          =   -1  'True
      TabIndex        =   131
      TabStop         =   0   'False
      Top             =   5520
      Width           =   465
   End
   Begin VB.TextBox Ra2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6675
      Locked          =   -1  'True
      TabIndex        =   130
      TabStop         =   0   'False
      Top             =   5040
      Width           =   465
   End
   Begin VB.TextBox Pu2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6675
      Locked          =   -1  'True
      TabIndex        =   129
      TabStop         =   0   'False
      Top             =   4560
      Width           =   465
   End
   Begin VB.TextBox P2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6675
      Locked          =   -1  'True
      TabIndex        =   128
      TabStop         =   0   'False
      Top             =   4080
      Width           =   465
   End
   Begin VB.TextBox Os2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6675
      Locked          =   -1  'True
      TabIndex        =   127
      TabStop         =   0   'False
      Top             =   3600
      Width           =   465
   End
   Begin VB.TextBox O2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6675
      Locked          =   -1  'True
      TabIndex        =   126
      TabStop         =   0   'False
      Top             =   3120
      Width           =   465
   End
   Begin VB.TextBox Np2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6675
      Locked          =   -1  'True
      TabIndex        =   125
      TabStop         =   0   'False
      Top             =   2640
      Width           =   465
   End
   Begin VB.TextBox No2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6675
      Locked          =   -1  'True
      TabIndex        =   124
      TabStop         =   0   'False
      Top             =   2160
      Width           =   465
   End
   Begin VB.TextBox N2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6675
      Locked          =   -1  'True
      TabIndex        =   123
      TabStop         =   0   'False
      Top             =   1680
      Width           =   465
   End
   Begin VB.TextBox Hg2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6675
      Locked          =   -1  'True
      TabIndex        =   122
      TabStop         =   0   'False
      Top             =   1200
      Width           =   465
   End
   Begin VB.TextBox Mn2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6675
      Locked          =   -1  'True
      TabIndex        =   121
      TabStop         =   0   'False
      Top             =   720
      Width           =   465
   End
   Begin VB.TextBox Mg2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4155
      Locked          =   -1  'True
      TabIndex        =   120
      TabStop         =   0   'False
      Top             =   6000
      Width           =   495
   End
   Begin VB.TextBox Li2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4155
      Locked          =   -1  'True
      TabIndex        =   119
      TabStop         =   0   'False
      Top             =   5520
      Width           =   495
   End
   Begin VB.TextBox Pb2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4155
      Locked          =   -1  'True
      TabIndex        =   118
      TabStop         =   0   'False
      Top             =   5040
      Width           =   495
   End
   Begin VB.TextBox Kr2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4155
      Locked          =   -1  'True
      TabIndex        =   117
      TabStop         =   0   'False
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox Fe2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4155
      Locked          =   -1  'True
      TabIndex        =   116
      TabStop         =   0   'False
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox I2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4155
      Locked          =   -1  'True
      TabIndex        =   115
      TabStop         =   0   'False
      Top             =   3600
      Width           =   495
   End
   Begin VB.TextBox H2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4155
      Locked          =   -1  'True
      TabIndex        =   114
      TabStop         =   0   'False
      Top             =   3120
      Width           =   495
   End
   Begin VB.TextBox Ga2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4155
      Locked          =   -1  'True
      TabIndex        =   113
      TabStop         =   0   'False
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox Au2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4155
      Locked          =   -1  'True
      TabIndex        =   112
      TabStop         =   0   'False
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox Fm2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4155
      Locked          =   -1  'True
      TabIndex        =   111
      TabStop         =   0   'False
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox F2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4155
      Locked          =   -1  'True
      TabIndex        =   110
      TabStop         =   0   'False
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox Es2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4155
      Locked          =   -1  'True
      TabIndex        =   109
      TabStop         =   0   'False
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox Co2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1755
      Locked          =   -1  'True
      TabIndex        =   108
      TabStop         =   0   'False
      Top             =   6000
      Width           =   465
   End
   Begin VB.TextBox Cu2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1755
      Locked          =   -1  'True
      TabIndex        =   107
      TabStop         =   0   'False
      Top             =   5520
      Width           =   465
   End
   Begin VB.TextBox Cl2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1755
      Locked          =   -1  'True
      TabIndex        =   106
      TabStop         =   0   'False
      Top             =   5040
      Width           =   465
   End
   Begin VB.TextBox C2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1755
      Locked          =   -1  'True
      TabIndex        =   104
      TabStop         =   0   'False
      Top             =   4560
      Width           =   465
   End
   Begin VB.TextBox Ca2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1755
      Locked          =   -1  'True
      TabIndex        =   102
      TabStop         =   0   'False
      Top             =   4080
      Width           =   465
   End
   Begin VB.TextBox Br2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1755
      Locked          =   -1  'True
      TabIndex        =   100
      TabStop         =   0   'False
      Top             =   3600
      Width           =   465
   End
   Begin VB.TextBox B2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1755
      Locked          =   -1  'True
      TabIndex        =   98
      TabStop         =   0   'False
      Top             =   3120
      Width           =   465
   End
   Begin VB.TextBox Be2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1755
      Locked          =   -1  'True
      TabIndex        =   96
      TabStop         =   0   'False
      Top             =   2640
      Width           =   465
   End
   Begin VB.TextBox Ba2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1755
      Locked          =   -1  'True
      TabIndex        =   94
      TabStop         =   0   'False
      Top             =   2160
      Width           =   465
   End
   Begin VB.TextBox txtAs2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1755
      Locked          =   -1  'True
      TabIndex        =   92
      TabStop         =   0   'False
      Top             =   1680
      Width           =   465
   End
   Begin VB.TextBox Sb2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1755
      Locked          =   -1  'True
      TabIndex        =   90
      TabStop         =   0   'False
      Top             =   1200
      Width           =   465
   End
   Begin VB.TextBox Al2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1755
      Locked          =   -1  'True
      TabIndex        =   88
      TabStop         =   0   'False
      Top             =   720
      Width           =   465
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Results"
      DownPicture     =   "Elements&Symbols.frx":CACA
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   120
      Picture         =   "Elements&Symbols.frx":E60C
      Style           =   1  'Graphical
      TabIndex        =   105
      Top             =   6600
      Width           =   7695
   End
   Begin VB.TextBox No 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6000
      MaxLength       =   2
      TabIndex        =   75
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox N 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6000
      MaxLength       =   2
      TabIndex        =   74
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox Hg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6000
      MaxLength       =   2
      TabIndex        =   73
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox Mn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6000
      MaxLength       =   2
      TabIndex        =   72
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox Mg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      MaxLength       =   2
      TabIndex        =   71
      Top             =   6000
      Width           =   615
   End
   Begin VB.TextBox Li 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      MaxLength       =   2
      TabIndex        =   70
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox Pb 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      MaxLength       =   2
      TabIndex        =   69
      Top             =   5040
      Width           =   615
   End
   Begin VB.TextBox Kr 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      MaxLength       =   2
      TabIndex        =   68
      Top             =   4560
      Width           =   615
   End
   Begin VB.TextBox Fe 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      MaxLength       =   2
      TabIndex        =   67
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox I 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      MaxLength       =   2
      TabIndex        =   66
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox H 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      MaxLength       =   2
      TabIndex        =   65
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Ga 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      MaxLength       =   2
      TabIndex        =   64
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox Au 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      MaxLength       =   2
      TabIndex        =   63
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox Fm 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      MaxLength       =   2
      TabIndex        =   62
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox F 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      MaxLength       =   2
      TabIndex        =   61
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox Es 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      MaxLength       =   2
      TabIndex        =   60
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox Co 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      MaxLength       =   2
      TabIndex        =   59
      Top             =   6000
      Width           =   615
   End
   Begin VB.TextBox Cu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      MaxLength       =   2
      TabIndex        =   58
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox Cl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      MaxLength       =   2
      TabIndex        =   57
      Top             =   5040
      Width           =   615
   End
   Begin VB.TextBox C 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      MaxLength       =   2
      TabIndex        =   56
      Top             =   4560
      Width           =   615
   End
   Begin VB.TextBox Ca 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      MaxLength       =   2
      TabIndex        =   55
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox Br 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      MaxLength       =   2
      TabIndex        =   54
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox B 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      MaxLength       =   2
      TabIndex        =   53
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Be 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      MaxLength       =   2
      TabIndex        =   52
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox Ba 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      MaxLength       =   2
      TabIndex        =   51
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox txtAs 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      MaxLength       =   2
      TabIndex        =   50
      ToolTipText     =   "Your Answer"
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox Sb 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      MaxLength       =   2
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   2  'Automatic
      TabIndex        =   49
      ToolTipText     =   "Your Answer"
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox Al 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      MaxLength       =   2
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   2  'Automatic
      TabIndex        =   48
      Tag             =   "Sb"
      ToolTipText     =   "Your Answer"
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox Zn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8400
      MaxLength       =   2
      TabIndex        =   103
      Top             =   6000
      Width           =   615
   End
   Begin VB.TextBox Yb 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8400
      MaxLength       =   2
      TabIndex        =   101
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox Xe 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8400
      MaxLength       =   2
      TabIndex        =   99
      Top             =   5040
      Width           =   615
   End
   Begin VB.TextBox Va 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8400
      MaxLength       =   2
      TabIndex        =   97
      Top             =   4560
      Width           =   615
   End
   Begin VB.TextBox U 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8400
      MaxLength       =   2
      TabIndex        =   95
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox Te 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8400
      MaxLength       =   2
      TabIndex        =   93
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox Ti 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8400
      MaxLength       =   2
      TabIndex        =   91
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Th 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8400
      MaxLength       =   2
      TabIndex        =   89
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox Sn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8400
      MaxLength       =   2
      TabIndex        =   87
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox S 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8400
      MaxLength       =   2
      TabIndex        =   86
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox Na 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8400
      MaxLength       =   2
      TabIndex        =   85
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox Sr 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8400
      MaxLength       =   2
      TabIndex        =   84
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox Ag 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6000
      MaxLength       =   2
      TabIndex        =   83
      Top             =   6000
      Width           =   615
   End
   Begin VB.TextBox Rn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6000
      MaxLength       =   2
      TabIndex        =   82
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox Ra 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6000
      MaxLength       =   2
      TabIndex        =   81
      Top             =   5040
      Width           =   615
   End
   Begin VB.TextBox Pu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6000
      MaxLength       =   2
      TabIndex        =   80
      Top             =   4560
      Width           =   615
   End
   Begin VB.TextBox P 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6000
      MaxLength       =   2
      TabIndex        =   79
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox Os 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6000
      MaxLength       =   2
      TabIndex        =   78
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox O 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6000
      MaxLength       =   2
      TabIndex        =   77
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Np 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6000
      MaxLength       =   2
      TabIndex        =   76
      Top             =   2640
      Width           =   615
   End
   Begin VB.Line Line5 
      X1              =   4800
      X2              =   4800
      Y1              =   600
      Y2              =   6480
   End
   Begin VB.Line Line4 
      X1              =   7200
      X2              =   7200
      Y1              =   600
      Y2              =   6480
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fill in the textboxes with their symbols"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   146
      Top             =   120
      Width           =   9375
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   11040
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line1 
      X1              =   2280
      X2              =   2280
      Y1              =   600
      Y2              =   6480
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Manganese"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   47
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Mercury"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   46
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Nitrogen"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   45
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Magnesium"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   44
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Lithium"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   43
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Lead"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   42
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Aluminium"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   41
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Antimony"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Arsenic"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Barium"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Lable5 
      BackStyle       =   0  'Transparent
      Caption         =   "Berylium"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   37
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Boron"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Bromine"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Calcium"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Carbon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "Nobelium"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   32
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "Neptunium"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   31
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "Oxygen"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   30
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "Osmium"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   29
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "Phosphorus"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   28
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label34 
      BackStyle       =   0  'Transparent
      Caption         =   "Plutonium"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   27
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label35 
      BackStyle       =   0  'Transparent
      Caption         =   "Radium"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   26
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label36 
      BackStyle       =   0  'Transparent
      Caption         =   "Radon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   25
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "Silver"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4920
      TabIndex        =   24
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label Label38 
      BackStyle       =   0  'Transparent
      Caption         =   "Stontium"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   23
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label39 
      BackStyle       =   0  'Transparent
      Caption         =   "Sodium"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   22
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label40 
      BackStyle       =   0  'Transparent
      Caption         =   "Sulphur"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   21
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label41 
      BackStyle       =   0  'Transparent
      Caption         =   "Tin"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   20
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label42 
      BackStyle       =   0  'Transparent
      Caption         =   "Thorium"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   19
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label43 
      BackStyle       =   0  'Transparent
      Caption         =   "Titanium"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   18
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label44 
      BackStyle       =   0  'Transparent
      Caption         =   "Tellurium"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   17
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label45 
      BackStyle       =   0  'Transparent
      Caption         =   "Uranium"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   16
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label46 
      BackStyle       =   0  'Transparent
      Caption         =   "Vanadium"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   15
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label47 
      BackStyle       =   0  'Transparent
      Caption         =   "Xenon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   14
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label48 
      BackStyle       =   0  'Transparent
      Caption         =   "Ytterbium"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   13
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label49 
      BackStyle       =   0  'Transparent
      Caption         =   "Zinc"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7320
      TabIndex        =   12
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Krypton"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   11
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Iron"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   10
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Iodine"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   9
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Hydrogen"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Gold"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Fermium"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Fluorine"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Einstenium"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Cobalt"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Copper"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Chlorine"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Galluim"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
End
Attribute VB_Name = "frmSymbols"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IntVariable As Integer

Private Sub Command1_Click()
If Al = "Al" Then
Al2 = "a"
Else
Al2 = "r"
End If
If Sb = "Sb" Then
Sb2 = "a"
Else: Sb2 = "r"
End If
If Ba = "Ba" Then
Ba2 = "a"
Else: Ba2 = "r"
End If
If txtAs = "As" Then
txtAs2 = "a"
Else: txtAs2 = "r"
End If
If Be = "Be" Then
Be2 = "a"
Else: Be2 = "r"
End If
If B = "B" Then
B2 = "a"
Else: B2 = "r"
End If
If Br = "Br" Then
Br2 = "a"
Else: Br2 = "r"
End If
If Ca = "Ca" Then
Ca2 = "a"
Else: Ca2 = "r"
End If
If C = "C" Then
C2 = "a"
Else: C2 = "r"
End If
If Cl = "Cl" Then
Cl2 = "a"
Else: Cl2 = "r"
End If
If Cu = "Cu" Then
Cu2 = "a"
Else: Cu2 = "r"
End If
If Co = "Co" Then
Co2 = "a"
Else: Co2 = "r"
End If
If Es = "Es" Then
Es2 = "a"
Else: Es2 = "r"
End If
If F = "F" Then
F2 = "a"
Else: F2 = "r"
End If
If Fm = "Fm" Then
Fm2 = "a"
Else: Fm2 = "r"
End If
If Au = "Au" Then
Au2 = "a"
Else: Au2 = "r"
End If
If Ga = "Ga" Then
Ga2 = "a"
Else: Ga2 = "r"
End If
If I = "I" Then
I2 = "a"
Else: I2 = "r"
End If
If H = "H" Then
H2 = "a"
Else: H2 = "r"
End If
If I = "I" Then
I2 = "a"
Else: I2 = "r"
End If
If Fe = "Fe" Then
Fe2 = "a"
Else: Fe2 = "r"
End If
If Kr = "Kr" Then
Kr2 = "a"
Else: Kr2 = "r"
End If
If Pb = "Pb" Then
Pb2 = "a"
Else: Pb2 = "r"
End If
If Li = "Li" Then
Li2 = "a"
Else: Li2 = "r"
End If
If Mg = "Mg" Then
Mg2 = "a"
Else: Mg2 = "r"
End If
If Mn = "Mn" Then
Mn2 = "a"
Else: Mn2 = "r"
End If
If Hg = "Hg" Then
Hg2 = "a"
Else: Hg2 = "r"
End If
If N = "N" Then
N2 = "a"
Else: N2 = "r"
End If
If No = "No" Then
No2 = "a"
Else: No2 = "r"
End If
If Np = "Np" Then
Np2 = "a"
Else: Np2 = "r"
End If
If O = "O" Then
O2 = "a"
Else: O2 = "r"
End If
If Os = "Os" Then
Os2 = "a"
Else: Os2 = "r"
End If
If P = "P" Then
P2 = "a"
Else: P2 = "r"
End If
If K = "K" Then
K2 = "a"
Else: K2 = "r"
End If
If Pu = "Pu" Then
Pu2 = "a"
Else: Pu2 = "r"
End If
If Ra = "Ra" Then
Ra2 = "a"
Else: Ra2 = "r"
End If
If Rn = "Rn" Then
Rn2 = "a"
Else: Rn2 = "r"
End If
If Ag = "Ag" Then
Ag2 = "a"
Else: Ag2 = "r"
End If
If Sr = "Sr" Then
Sr2 = "a"
Else: Sr2 = "r"
End If
If Na = "Na" Then
Na2 = "a"
Else: Na2 = "r"
End If
If S = "S" Then
S2 = "a"
Else: S2 = "r"
End If
If Sn = "Sn" Then
Sn2 = "a"
Else: Sn2 = "r"
End If
If Th = "Th" Then
Th2 = "a"
Else: Th2 = "r"
End If
If Ti = "Ti" Then
Ti2 = "a"
Else: Ti2 = "r"
End If
If Te = "Te" Then
Te2 = "a"
Else: Te2 = "r"
End If
If U = "U" Then
U2 = "a"
Else: U2 = "r"
End If
If Va = "Va" Then
Va2 = "a"
Else: Va2 = "r"
End If
If Xe = "Xe" Then
Xe2 = "a"
Else: Xe2 = "r"
End If
If Yb = "Yb" Then
Yb2 = "a"
Else: Yb2 = "r"
End If
If Y = "Y" Then
Y2 = "a"
Else: Y2 = "r"
End If
If Zn = "Zn" Then
Zn2 = "a"
Else: Zn2 = "r"
End If
If Zr = "Zr" Then
Zr2 = "a"
Else: Zr2 = "r"
End If
End Sub

Private Sub Command2_Click()
Me.Hide
frmWelcomescreen.Show
End Sub

Private Sub Form_Load()
IntVariable = MsgBox("Welcome to Chemistry Quiz. Lets test your knowledge on symbols! Best of luck!!", vbOKOnly, "Welcome")
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = True
Me.Hide
frmWelcomescreen.Show
End Sub

