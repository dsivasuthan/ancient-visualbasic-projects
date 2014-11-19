VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unit Converter - Dhayalan Sivasuthan"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5460
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   5460
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "About"
      Height          =   375
      Left            =   3795
      TabIndex        =   38
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Output"
      Height          =   615
      Left            =   120
      TabIndex        =   34
      Top             =   4080
      Width           =   3495
      Begin VB.TextBox txt2ndUnit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lbl2ndUnit 
         Height          =   255
         Left            =   1800
         TabIndex        =   36
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Input"
      Height          =   615
      Left            =   120
      TabIndex        =   32
      Top             =   3240
      Width           =   3495
      Begin VB.TextBox txt1stUnit 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         TabIndex        =   33
         Text            =   "1"
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lbl1stUnit 
         Height          =   255
         Left            =   1800
         TabIndex        =   37
         Top             =   240
         Width           =   1575
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   5318
      _Version        =   393216
      Style           =   1
      Tabs            =   10
      TabsPerRow      =   5
      TabHeight       =   520
      TabMaxWidth     =   4
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Length"
      TabPicture(0)   =   "Form1.frx":1CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Length(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Length(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Area"
      TabPicture(1)   =   "Form1.frx":1CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2(1)"
      Tab(1).Control(1)=   "Area(0)"
      Tab(1).Control(2)=   "Area(1)"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Mass"
      TabPicture(2)   =   "Form1.frx":1D02
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Mass(0)"
      Tab(2).Control(1)=   "Mass(1)"
      Tab(2).Control(2)=   "Label2(2)"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Volume"
      TabPicture(3)   =   "Form1.frx":1D1E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label2(3)"
      Tab(3).Control(1)=   "Volume(1)"
      Tab(3).Control(2)=   "Volume(0)"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Density"
      TabPicture(4)   =   "Form1.frx":1D3A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label2(4)"
      Tab(4).Control(1)=   "Density(0)"
      Tab(4).Control(2)=   "Density(1)"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "Energy"
      TabPicture(5)   =   "Form1.frx":1D56
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label2(5)"
      Tab(5).Control(1)=   "Energy(0)"
      Tab(5).Control(2)=   "Energy(1)"
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "Power"
      TabPicture(6)   =   "Form1.frx":1D72
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Power(0)"
      Tab(6).Control(1)=   "Power(1)"
      Tab(6).Control(2)=   "Label2(6)"
      Tab(6).ControlCount=   3
      TabCaption(7)   =   "Pressure"
      TabPicture(7)   =   "Form1.frx":1D8E
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Pressure(0)"
      Tab(7).Control(1)=   "Pressure(1)"
      Tab(7).Control(2)=   "Label2(7)"
      Tab(7).ControlCount=   3
      TabCaption(8)   =   "Speed"
      TabPicture(8)   =   "Form1.frx":1DAA
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Label2(8)"
      Tab(8).Control(1)=   "Speed(0)"
      Tab(8).Control(2)=   "Speed(1)"
      Tab(8).ControlCount=   3
      TabCaption(9)   =   "Temperature"
      TabPicture(9)   =   "Form1.frx":1DC6
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Temp(0)"
      Tab(9).Control(1)=   "Temp(1)"
      Tab(9).Control(2)=   "Label2(9)"
      Tab(9).ControlCount=   3
      Begin VB.ComboBox Temp 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Index           =   0
         ItemData        =   "Form1.frx":1DE2
         Left            =   -74880
         List            =   "Form1.frx":1DEF
         Style           =   1  'Simple Combo
         TabIndex        =   21
         Text            =   "kelvin"
         Top             =   840
         Width           =   2055
      End
      Begin VB.ComboBox Temp 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Index           =   1
         ItemData        =   "Form1.frx":1E0F
         Left            =   -72120
         List            =   "Form1.frx":1E1C
         Style           =   1  'Simple Combo
         TabIndex        =   20
         Text            =   "Celsius"
         Top             =   840
         Width           =   2055
      End
      Begin VB.ComboBox Speed 
         Height          =   1740
         Index           =   1
         ItemData        =   "Form1.frx":1E3C
         Left            =   -72240
         List            =   "Form1.frx":1E49
         Style           =   1  'Simple Combo
         TabIndex        =   19
         Text            =   "m/sec"
         Top             =   840
         Width           =   1935
      End
      Begin VB.ComboBox Speed 
         Height          =   1740
         Index           =   0
         ItemData        =   "Form1.frx":1E64
         Left            =   -74880
         List            =   "Form1.frx":1E71
         Style           =   1  'Simple Combo
         TabIndex        =   18
         Text            =   "km/hr"
         Top             =   840
         Width           =   1935
      End
      Begin VB.ComboBox Pressure 
         Height          =   1740
         Index           =   0
         ItemData        =   "Form1.frx":1E8C
         Left            =   -74880
         List            =   "Form1.frx":1E99
         Style           =   1  'Simple Combo
         TabIndex        =   17
         Text            =   "pascal"
         Top             =   840
         Width           =   1935
      End
      Begin VB.ComboBox Pressure 
         Height          =   1740
         Index           =   1
         ItemData        =   "Form1.frx":1EB7
         Left            =   -72240
         List            =   "Form1.frx":1EC4
         Style           =   1  'Simple Combo
         TabIndex        =   16
         Text            =   "mmHg"
         Top             =   840
         Width           =   1935
      End
      Begin VB.ComboBox Power 
         Height          =   1740
         Index           =   0
         ItemData        =   "Form1.frx":1EE2
         Left            =   -74880
         List            =   "Form1.frx":1EEF
         Style           =   1  'Simple Combo
         TabIndex        =   15
         Text            =   "watt"
         Top             =   840
         Width           =   1935
      End
      Begin VB.ComboBox Power 
         Height          =   1740
         Index           =   1
         ItemData        =   "Form1.frx":1F0F
         Left            =   -72240
         List            =   "Form1.frx":1F1C
         Style           =   1  'Simple Combo
         TabIndex        =   14
         Text            =   "hp (metric)"
         Top             =   840
         Width           =   1935
      End
      Begin VB.ComboBox Energy 
         Height          =   1740
         Index           =   1
         ItemData        =   "Form1.frx":1F3C
         Left            =   -72120
         List            =   "Form1.frx":1F4C
         Style           =   1  'Simple Combo
         TabIndex        =   13
         Text            =   "cal"
         Top             =   720
         Width           =   1935
      End
      Begin VB.ComboBox Energy 
         Height          =   1740
         Index           =   0
         ItemData        =   "Form1.frx":1F66
         Left            =   -74880
         List            =   "Form1.frx":1F76
         Style           =   1  'Simple Combo
         TabIndex        =   12
         Text            =   "joule"
         Top             =   720
         Width           =   1935
      End
      Begin VB.ComboBox Density 
         Height          =   1935
         Index           =   1
         ItemData        =   "Form1.frx":1F90
         Left            =   -72240
         List            =   "Form1.frx":1F9D
         Style           =   1  'Simple Combo
         TabIndex        =   11
         Text            =   "gram/cu cm"
         Top             =   720
         Width           =   1935
      End
      Begin VB.ComboBox Density 
         Height          =   1935
         Index           =   0
         ItemData        =   "Form1.frx":1FC6
         Left            =   -74880
         List            =   "Form1.frx":1FD3
         Style           =   1  'Simple Combo
         TabIndex        =   10
         Text            =   "lb/cu metre"
         Top             =   720
         Width           =   1935
      End
      Begin VB.ComboBox Volume 
         Height          =   1935
         Index           =   0
         ItemData        =   "Form1.frx":1FFC
         Left            =   -74880
         List            =   "Form1.frx":2012
         Style           =   1  'Simple Combo
         TabIndex        =   9
         Text            =   "litre (cu dm)"
         Top             =   840
         Width           =   1935
      End
      Begin VB.ComboBox Volume 
         Height          =   1935
         Index           =   1
         ItemData        =   "Form1.frx":205B
         Left            =   -72240
         List            =   "Form1.frx":2071
         Style           =   1  'Simple Combo
         TabIndex        =   8
         Text            =   "cu metre"
         Top             =   840
         Width           =   1935
      End
      Begin VB.ComboBox Mass 
         BeginProperty Font 
            Name            =   "Verdana Ref"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1650
         Index           =   0
         ItemData        =   "Form1.frx":20BA
         Left            =   -74880
         List            =   "Form1.frx":20CA
         Style           =   1  'Simple Combo
         TabIndex        =   7
         Text            =   "ounce"
         Top             =   840
         Width           =   1935
      End
      Begin VB.ComboBox Mass 
         BeginProperty Font 
            Name            =   "Verdana Ref"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1650
         Index           =   1
         ItemData        =   "Form1.frx":20E9
         Left            =   -72240
         List            =   "Form1.frx":20F9
         Style           =   1  'Simple Combo
         TabIndex        =   6
         Text            =   "gram"
         Top             =   840
         Width           =   1935
      End
      Begin VB.ComboBox Area 
         Height          =   1740
         Index           =   1
         ItemData        =   "Form1.frx":2118
         Left            =   -72240
         List            =   "Form1.frx":2131
         Style           =   1  'Simple Combo
         TabIndex        =   5
         Text            =   "square kilometre"
         Top             =   780
         Width           =   1935
      End
      Begin VB.ComboBox Area 
         Height          =   1740
         Index           =   0
         ItemData        =   "Form1.frx":2191
         Left            =   -74880
         List            =   "Form1.frx":21AA
         Style           =   1  'Simple Combo
         TabIndex        =   4
         Text            =   "acre"
         Top             =   780
         Width           =   1935
      End
      Begin VB.ComboBox Length 
         Height          =   1935
         Index           =   0
         ItemData        =   "Form1.frx":220A
         Left            =   120
         List            =   "Form1.frx":222F
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   2
         Text            =   "metre"
         Top             =   780
         Width           =   1935
      End
      Begin VB.ComboBox Length 
         Height          =   1935
         Index           =   1
         ItemData        =   "Form1.frx":2297
         Left            =   2760
         List            =   "Form1.frx":22BC
         Style           =   1  'Simple Combo
         TabIndex        =   1
         Text            =   "inch"
         Top             =   780
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "->"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   9
         Left            =   -72720
         TabIndex        =   31
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "->"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   8
         Left            =   -72840
         TabIndex        =   30
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "->"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   -72840
         TabIndex        =   29
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "->"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   -72840
         TabIndex        =   28
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "->"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   -72840
         TabIndex        =   27
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "->"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   -72840
         TabIndex        =   26
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "->"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   -72840
         TabIndex        =   25
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "->"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   -72840
         TabIndex        =   24
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "->"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   -72840
         TabIndex        =   23
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "->"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   2160
         TabIndex        =   22
         Top             =   1440
         Width           =   1095
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "="
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   3840
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public DEx As Boolean


Private Sub Area_Click(Index As Integer)
lbl1stUnit.Caption = Area(0).Text
lbl2ndUnit.Caption = Area(1).Text


If Area(0).Text = "square metre" Then
If Area(1).Text = "square metre" Then
 txt2ndUnit = txt1stUnit * 1
Else
If Area(1).Text = "hectare" Then
  txt2ndUnit = txt1stUnit * 0.0001
Else
If Area(1).Text = "acre" Then
  txt2ndUnit = txt1stUnit * 0.000247
Else
If Area(1).Text = "square mile" Then
  txt2ndUnit = txt1stUnit * 0.0000003861022
Else
If Area(1).Text = "square kilometre" Then
  txt2ndUnit = txt1stUnit * 0.000001

End If
End If
End If
End If
End If
End If


If Area(0).Text = "square kilometre" Then
If Area(1).Text = "square metre" Then
 txt2ndUnit = txt1stUnit * 1000000
Else
If Area(1).Text = "hectare" Then
  txt2ndUnit = txt1stUnit * 100
Else
If Area(1).Text = "acre" Then
  txt2ndUnit = txt1stUnit * 247.105381
Else
If Area(1).Text = "square mile" Then
  txt2ndUnit = txt1stUnit * 0.386102
Else
If Area(1).Text = "square kilometre" Then
  txt2ndUnit = txt1stUnit * 1

End If
End If
End If
End If
End If
End If



If Area(0).Text = "hectare" Then
If Area(1).Text = "square metre" Then
 txt2ndUnit = txt1stUnit * 10000
Else
If Area(1).Text = "hectare" Then
  txt2ndUnit = txt1stUnit * 1
Else
If Area(1).Text = "acre" Then
  txt2ndUnit = txt1stUnit * 2.471054
Else
If Area(1).Text = "square mile" Then
  txt2ndUnit = txt1stUnit * 0.003861
Else
If Area(1).Text = "square kilometre" Then
  txt2ndUnit = txt1stUnit * 0.01

End If
End If
End If
End If
End If
End If


If Area(0).Text = "acre" Then
If Area(1).Text = "square metre" Then
 txt2ndUnit = txt1stUnit * 4046.856422
Else
If Area(1).Text = "hectare" Then
  txt2ndUnit = txt1stUnit * 0.404686
Else
If Area(1).Text = "acre" Then
  txt2ndUnit = txt1stUnit * 1
Else
If Area(1).Text = "square mile" Then
  txt2ndUnit = txt1stUnit * 0.001563
Else
If Area(1).Text = "square kilometre" Then
  txt2ndUnit = txt1stUnit * 0.004047

End If
End If
End If
End If
End If
End If


If Area(0).Text = "square mile" Then
If Area(1).Text = "square metre" Then
 txt2ndUnit = txt1stUnit * 2589988.110336
Else
If Area(1).Text = "hectare" Then
  txt2ndUnit = txt1stUnit * 258.998811
Else
If Area(1).Text = "acre" Then
  txt2ndUnit = txt1stUnit * 640
Else
If Area(1).Text = "square mile" Then
  txt2ndUnit = txt1stUnit * 1
Else
If Area(1).Text = "square kilometre" Then
  txt2ndUnit = txt1stUnit * 2.589988

End If
End If
End If
End If
End If
End If

If Area(0).Text = "square inch" Then
If Area(1).Text = "square metre" Then txt2ndUnit = txt1stUnit * 0.00064516
If Area(1).Text = "hectare" Then txt2ndUnit = txt1stUnit * 0.000000064516
If Area(1).Text = "acre" Then txt2ndUnit = txt1stUnit * 0.0000001594219
If Area(1).Text = "square mile" Then txt2ndUnit = txt1stUnit * 2.490977E-10
If Area(1).Text = "square kilometre" Then txt2ndUnit = txt1stUnit * 0.00000000064516
If Area(1).Text = "square centimetre" Then txt2ndUnit = txt1stUnit * 6.4516
If Area(1).Text = "square inch" Then txt2ndUnit = txt1stUnit * 1

End If





End Sub

Private Sub Command1_Click()
Me.Hide
Dialog.Show
End Sub

Private Sub Density_Click(Index As Integer)
lbl1stUnit.Caption = Density(0).Text
lbl2ndUnit.Caption = Density(1).Text
If Density(0).Text = "kg/cu metre" Then
If Density(1).Text = "kg/cu metre" Then
 txt2ndUnit = txt1stUnit * 1
Else
If Density(1).Text = "gram/cu cm" Then
  txt2ndUnit = txt1stUnit * 0.001
Else
If Density(1).Text = "lb/cu inch" Then
  txt2ndUnit = txt1stUnit * 0.000036
End If
End If
End If
End If



If Density(0).Text = "gram/cu cm" Then
If Density(1).Text = "kg/cu metre" Then
 txt2ndUnit = txt1stUnit * 1000
Else
If Density(1).Text = "gram/cu cm" Then
  txt2ndUnit = txt1stUnit * 1
Else
If Density(1).Text = "lb/cu inch" Then
  txt2ndUnit = txt1stUnit * 0.036127
End If
End If
End If
End If


If Density(0).Text = "lb/cu inch" Then
If Density(1).Text = "kg/cu metre" Then
 txt2ndUnit = txt1stUnit * 27679.90498
Else
If Density(1).Text = "gram/cu cm" Then
  txt2ndUnit = txt1stUnit * 27.679905
Else
If Density(1).Text = "lb/cu inch" Then
  txt2ndUnit = txt1stUnit * 1
End If
End If
End If
End If


End Sub

Private Sub Energy_Click(Index As Integer)
lbl1stUnit.Caption = Energy(0).Text
lbl2ndUnit.Caption = Energy(1).Text
If Energy(0).Text = "joule" Then
If Energy(1).Text = "joule" Then
 txt2ndUnit = txt1stUnit * 1
Else
If Energy(1).Text = "cal" Then
  txt2ndUnit = txt1stUnit * 0.238846
Else
If Energy(1).Text = "kWh" Then
  txt2ndUnit = txt1stUnit * 2.77777777777778E-07
Else
If Energy(1).Text = "erg" Then
  txt2ndUnit = txt1stUnit * 10000000
End If
End If
End If
End If
End If



If Energy(0).Text = "cal" Then
If Energy(1).Text = "joule" Then
 txt2ndUnit = txt1stUnit * 4.1868
Else
If Energy(1).Text = "cal" Then
  txt2ndUnit = txt1stUnit * 1
Else
If Energy(1).Text = "kWh" Then
  txt2ndUnit = txt1stUnit * 0.000001163
Else
If Energy(1).Text = "erg" Then
  txt2ndUnit = txt1stUnit * 41868000
End If
End If
End If
End If
End If



If Energy(0).Text = "kWh" Then
If Energy(1).Text = "joule" Then
 txt2ndUnit = txt1stUnit * 3600000
Else
If Energy(1).Text = "cal" Then
  txt2ndUnit = txt1stUnit * 859845.227859
Else
If Energy(1).Text = "kWh" Then
  txt2ndUnit = txt1stUnit * 1
Else
If Energy(1).Text = "erg" Then
  txt2ndUnit = txt1stUnit * 36000000000000#
End If
End If
End If
End If
End If



If Energy(0).Text = "erg" Then
If Energy(1).Text = "joule" Then
 txt2ndUnit = txt1stUnit * 0.0000001
Else
If Energy(1).Text = "cal" Then
  txt2ndUnit = txt1stUnit * 0.0000000238846
Else
If Energy(1).Text = "kWh" Then
  txt2ndUnit = txt1stUnit * 1
Else
If Energy(1).Text = "erg" Then
  txt2ndUnit = txt1stUnit * 2.7777778E-14
End If
End If
End If
End If
End If


End Sub

Private Sub Form_Unload(Cancel As Integer)
'DEx = True
'Me.Hide
'Dialog.Show
End
End Sub

Private Sub Length_Click(Index As Integer)
lbl1stUnit.Caption = Length(0).Text
lbl2ndUnit.Caption = Length(1).Text
If Length(0).Text = "millimetre" Then
If Length(1).Text = "millimetre" Then
 txt2ndUnit = txt1stUnit / 1
Else
If Length(1).Text = "centimetre" Then
  txt2ndUnit = txt1stUnit / 10
Else
If Length(1).Text = "decimetre" Then
  txt2ndUnit = txt1stUnit / 100
Else
If Length(1).Text = "metre" Then
    txt2ndUnit = txt1stUnit / 1000
Else
If Length(1).Text = "decametre" Then
    txt2ndUnit = txt1stUnit / 10000
Else
If Length(1).Text = "hectometre" Then
    txt2ndUnit = txt1stUnit / 100000
Else
If Length(1).Text = "kilometre" Then
    txt2ndUnit = txt1stUnit / 1000000
Else
If Length(1).Text = "inch" Then
    txt2ndUnit = txt1stUnit * 0.03937
Else
If Length(1).Text = "foot" Then
    txt2ndUnit = txt1stUnit * 0.003281
Else
If Length(1).Text = "yard" Then
    txt2ndUnit = txt1stUnit * 0.001094
Else
If Length(1).Text = "mile" Then
    txt2ndUnit = txt1stUnit * 0.000001

End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If




If Length(0).Text = "centimetre" Then
If Length(1).Text = "millimetre" Then
 txt2ndUnit = txt1stUnit * 10
Else
If Length(1).Text = "centimetre" Then
  txt2ndUnit = txt1stUnit / 1
Else
If Length(1).Text = "decimetre" Then
  txt2ndUnit = txt1stUnit / 10
Else
If Length(1).Text = "metre" Then
    txt2ndUnit = txt1stUnit / 100
Else
If Length(1).Text = "decametre" Then
    txt2ndUnit = txt1stUnit / 1000
Else
If Length(1).Text = "hectometre" Then
    txt2ndUnit = txt1stUnit / 10000
Else
If Length(1).Text = "kilometre" Then
    txt2ndUnit = txt1stUnit / 100000
Else
If Length(1).Text = "inch" Then
    txt2ndUnit = txt1stUnit * 0.3937
Else
If Length(1).Text = "foot" Then
    txt2ndUnit = txt1stUnit * 0.032808
Else
If Length(1).Text = "yard" Then
    txt2ndUnit = txt1stUnit * 0.0010936
Else
If Length(1).Text = "mile" Then
    txt2ndUnit = txt1stUnit * 0.000006


End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If


If Length(0).Text = "decimetre" Then
If Length(1).Text = "millimetre" Then
 txt2ndUnit = txt1stUnit * 100
Else
If Length(1).Text = "centimetre" Then
  txt2ndUnit = txt1stUnit * 10
Else
If Length(1).Text = "decimetre" Then
  txt2ndUnit = txt1stUnit / 1
Else
If Length(1).Text = "metre" Then
    txt2ndUnit = txt1stUnit / 10
Else
If Length(1).Text = "decametre" Then
    txt2ndUnit = txt1stUnit / 100
Else
If Length(1).Text = "hectometre" Then
    txt2ndUnit = txt1stUnit / 1000
Else
If Length(1).Text = "kilometre" Then
    txt2ndUnit = txt1stUnit / 10000
Else
If Length(1).Text = "inch" Then
    txt2ndUnit = txt1stUnit * 3.937008
Else
If Length(1).Text = "foot" Then
    txt2ndUnit = txt1stUnit * 0.328084
Else
If Length(1).Text = "yard" Then
    txt2ndUnit = txt1stUnit * 0.109361
Else
If Length(1).Text = "mile" Then
    txt2ndUnit = txt1stUnit * 0.000062


End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If


If Length(0).Text = "metre" Then
If Length(1).Text = "millimetre" Then
 txt2ndUnit = txt1stUnit * 1000
Else
If Length(1).Text = "centimetre" Then
  txt2ndUnit = txt1stUnit * 100
Else
If Length(1).Text = "decimetre" Then
  txt2ndUnit = txt1stUnit * 10
Else
If Length(1).Text = "metre" Then
    txt2ndUnit = txt1stUnit / 1
Else
If Length(1).Text = "decametre" Then
    txt2ndUnit = txt1stUnit / 10
Else
If Length(1).Text = "hectometre" Then
    txt2ndUnit = txt1stUnit / 100
Else
If Length(1).Text = "kilometre" Then
    txt2ndUnit = txt1stUnit / 1000
Else
If Length(1).Text = "inch" Then
    txt2ndUnit = txt1stUnit * 39.370079
Else
If Length(1).Text = "foot" Then
    txt2ndUnit = txt1stUnit * 3.28084
Else
If Length(1).Text = "yard" Then
    txt2ndUnit = txt1stUnit * 1.093613
Else
If Length(1).Text = "mile" Then
    txt2ndUnit = txt1stUnit * 0.000621


End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If


If Length(0).Text = "decametre" Then
If Length(1).Text = "millimetre" Then
 txt2ndUnit = txt1stUnit * 10000
Else
If Length(1).Text = "centimetre" Then
  txt2ndUnit = txt1stUnit * 1000
Else
If Length(1).Text = "decimetre" Then
  txt2ndUnit = txt1stUnit * 100
Else
If Length(1).Text = "metre" Then
    txt2ndUnit = txt1stUnit * 10
Else
If Length(1).Text = "decametre" Then
    txt2ndUnit = txt1stUnit / 1
Else
If Length(1).Text = "hectometre" Then
    txt2ndUnit = txt1stUnit / 10
Else
If Length(1).Text = "kilometre" Then
    txt2ndUnit = txt1stUnit / 100
Else
If Length(1).Text = "inch" Then
    txt2ndUnit = txt1stUnit * 393.700787
Else
If Length(1).Text = "foot" Then
    txt2ndUnit = txt1stUnit * 32.808399
Else
If Length(1).Text = "yard" Then
    txt2ndUnit = txt1stUnit * 10.936133
Else
If Length(1).Text = "mile" Then
    txt2ndUnit = txt1stUnit * 0.006214


End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If


If Length(0).Text = "inch" Then
If Length(1).Text = "centimetre" Then txt2ndUnit = txt1stUnit * 2.54
End If

If Length(0).Text = "foot" Then
If Length(1).Text = "metre" Then txt2ndUnit = txt1stUnit * 0.3048
End If

If Length(0).Text = "yard" Then
If Length(1).Text = "metre" Then txt2ndUnit = txt1stUnit * 0.9144
End If

If Length(0).Text = "mile" Then
If Length(1).Text = "kilometre" Then txt2ndUnit = txt1stUnit * 1.60934
End If


End Sub

Private Sub Mass_Click(Index As Integer)
lbl1stUnit.Caption = Mass(0).Text
lbl2ndUnit.Caption = Mass(1).Text
If Mass(0).Text = "gram" Then
If Mass(1).Text = "gram" Then
 txt2ndUnit = txt1stUnit * 1
Else
If Mass(1).Text = "kilogram" Then
  txt2ndUnit = txt1stUnit / 1000
Else
If Mass(1).Text = "ounce" Then
  txt2ndUnit = txt1stUnit * 0.035273962
Else
If Mass(1).Text = "lb" Then
    txt2ndUnit = txt1stUnit * 0.002204623

End If
End If
End If
End If
End If



If Mass(0).Text = "kilogram" Then
If Mass(1).Text = "gram" Then
 txt2ndUnit = txt1stUnit * 1000
Else
If Mass(1).Text = "kilogram" Then
  txt2ndUnit = txt1stUnit * 1
Else
If Mass(1).Text = "ounce" Then
  txt2ndUnit = txt1stUnit * 35.273962
Else
If Mass(1).Text = "lb" Then
    txt2ndUnit = txt1stUnit * 2.204623

End If
End If
End If
End If
End If



If Mass(0).Text = "ounce" Then
If Mass(1).Text = "gram" Then
 txt2ndUnit = txt1stUnit * 28.349523
Else
If Mass(1).Text = "kilogram" Then
  txt2ndUnit = txt1stUnit * 0.02835
Else
If Mass(1).Text = "ounce" Then
  txt2ndUnit = txt1stUnit * 1
Else
If Mass(1).Text = "lb" Then
    txt2ndUnit = txt1stUnit * 0.0625

End If
End If
End If
End If
End If



If Mass(0).Text = "lb" Then
If Mass(1).Text = "gram" Then
 txt2ndUnit = txt1stUnit * 453.592374
Else
If Mass(1).Text = "kilogram" Then
  txt2ndUnit = txt1stUnit * 0.453592
Else
If Mass(1).Text = "ounce" Then
  txt2ndUnit = txt1stUnit * 16
Else
If Mass(1).Text = "lb" Then
    txt2ndUnit = txt1stUnit * 1

End If
End If
End If
End If
End If

End Sub

Private Sub Power_Click(Index As Integer)
lbl1stUnit.Caption = Power(0).Text
lbl2ndUnit.Caption = Power(1).Text
If Power(0).Text = "watt" Then
If Power(1).Text = "watt" Then
txt2ndUnit = txt1stUnit * 1
Else
If Power(1).Text = "hp (metric)" Then
txt2ndUnit = txt1stUnit * 0.00136
Else
If Power(1).Text = "hp (UK)" Then
txt2ndUnit = txt1stUnit * 0.001341

End If
End If
End If
End If



If Power(0).Text = "hp (metric)" Then
If Power(1).Text = "watt" Then
txt2ndUnit = txt1stUnit * 735.49875
Else
If Power(1).Text = "hp (metric)" Then
txt2ndUnit = txt1stUnit * 1
Else
If Power(1).Text = "hp (UK)" Then
txt2ndUnit = txt1stUnit * 0.98632

End If
End If
End If
End If


If Power(0).Text = "hp (UK)" Then
If Power(1).Text = "watt" Then
txt2ndUnit = txt1stUnit * 745.699871
Else
If Power(1).Text = "hp (metric)" Then
txt2ndUnit = txt1stUnit * 1.01387
Else
If Power(1).Text = "hp (UK)" Then
txt2ndUnit = txt1stUnit * 1

End If
End If
End If
End If


End Sub

Private Sub Pressure_Click(Index As Integer)
lbl1stUnit.Caption = Pressure(0).Text
lbl2ndUnit.Caption = Pressure(1).Text
If Pressure(0).Text = "pascal" Then
If Pressure(1).Text = "pascal" Then
txt2ndUnit = txt1stUnit * 1
Else
If Pressure(1).Text = "mmHg" Then
txt2ndUnit = txt1stUnit * 0.007501
Else
If Pressure(1).Text = "atmosphere" Then
txt2ndUnit = txt1stUnit * 0.00001

End If
End If
End If
End If



If Pressure(0).Text = "atmosphere" Then
If Pressure(1).Text = "pascal" Then
txt2ndUnit = txt1stUnit * 101325
Else
If Pressure(1).Text = "atmosphere" Then
txt2ndUnit = txt1stUnit * 1
Else
If Pressure(1).Text = "mmHg" Then
txt2ndUnit = txt1stUnit * 759.999892
End If
End If
End If
End If



If Pressure(0).Text = "mmHg" Then
If Pressure(1).Text = "pascal" Then
txt2ndUnit = txt1stUnit * 133.322387
Else
If Pressure(1).Text = "atmosphere" Then
txt2ndUnit = txt1stUnit * 0.001316
Else
If Pressure(1).Text = "mmHg" Then
txt2ndUnit = txt1stUnit * 1
End If
End If
End If
End If


End Sub

Private Sub Speed_Click(Index As Integer)
lbl1stUnit.Caption = Speed(0).Text
lbl2ndUnit.Caption = Speed(1).Text
If Speed(0).Text = "km/hr" Then
If Speed(1).Text = "km/hr" Then
txt2ndUnit = txt1stUnit * 1
Else
If Speed(1).Text = "m/sec" Then
txt2ndUnit = (txt1stUnit / 3600) * 1000
Else
If Speed(1).Text = "mile/hr" Then
txt2ndUnit = (txt1stUnit / 1.609) * 1
End If
End If
End If
End If


If Speed(0).Text = "m/sec" Then
If Speed(1).Text = "m/sec" Then
txt2ndUnit = txt1stUnit * 1
Else
If Speed(1).Text = "km/hr" Then
txt2ndUnit = (txt1stUnit / 1000) * 3600
Else
If Speed(1).Text = "mile/hr" Then
txt2ndUnit = (txt1stUnit / 1609.344) * 3600
End If
End If
End If
End If


If Speed(0).Text = "mile/hr" Then
If Speed(1).Text = "mile/hr" Then
txt2ndUnit = txt1stUnit * 1
Else
If Speed(1).Text = "km/hr" Then
txt2ndUnit = (txt1stUnit / 1) * 1.609344
Else
If Speed(1).Text = "m/sec" Then
txt2ndUnit = (txt1stUnit * 3600) * 1609.344
End If
End If
End If
End If


End Sub



Private Sub Volume_Click(Index As Integer)
lbl1stUnit.Caption = Volume(0).Text
lbl2ndUnit.Caption = Volume(1).Text
If Volume(0) = "litre (cu dm)" Then
If Volume(1) = "litre (cu dm)" Then
 txt2ndUnit = txt1stUnit * 1
Else
If Volume(1) = "cu metre" Then
  txt2ndUnit = txt1stUnit * 0.001
Else
If Volume(1) = "cu inch" Then
  txt2ndUnit = txt1stUnit * 61.023744
Else
If Volume(1) = "cu foot" Then
  txt2ndUnit = txt1stUnit * 0.035315
Else
If Volume(1) = "gallon (UK)" Then
  txt2ndUnit = txt1stUnit * 0.219969
Else
If Volume(1) = "gallon (US)" Then
  txt2ndUnit = txt1stUnit * 0.264172
End If
End If
End If
End If
End If
End If
End If



If Volume(0) = "cu metre" Then
If Volume(1) = "litre (cu dm)" Then
 txt2ndUnit = txt1stUnit * 1000
Else
If Volume(1) = "cu metre" Then
  txt2ndUnit = txt1stUnit * 1
Else
If Volume(1) = "cu inch" Then
  txt2ndUnit = txt1stUnit * 61023.744095
Else
If Volume(1) = "cu foot" Then
  txt2ndUnit = txt1stUnit * 35.314667
Else
If Volume(1) = "gallon (UK)" Then
  txt2ndUnit = txt1stUnit * 219.969248
Else
If Volume(1) = "gallon (US)" Then
  txt2ndUnit = txt1stUnit * 264.172052
End If
End If
End If
End If
End If
End If
End If


If Volume(0) = "cu inch" Then
If Volume(1) = "litre (cu dm)" Then
 txt2ndUnit = txt1stUnit * 0.016387
Else
If Volume(1) = "cu metre" Then
  txt2ndUnit = txt1stUnit * 0.000016
Else
If Volume(1) = "cu inch" Then
  txt2ndUnit = txt1stUnit * 1
Else
If Volume(1) = "cu foot" Then
  txt2ndUnit = txt1stUnit * 0.000579
Else
If Volume(1) = "gallon (UK)" Then
  txt2ndUnit = txt1stUnit * 0.003605
Else
If Volume(1) = "gallon (US)" Then
  txt2ndUnit = txt1stUnit * 0.004329
End If
End If
End If
End If
End If
End If
End If


If Volume(0) = "cu foot" Then
If Volume(1) = "litre (cu dm)" Then
 txt2ndUnit = txt1stUnit * 28.316847
Else
If Volume(1) = "cu metre" Then
  txt2ndUnit = txt1stUnit * 0.028317
Else
If Volume(1) = "cu inch" Then
  txt2ndUnit = txt1stUnit * 1728
Else
If Volume(1) = "cu foot" Then
  txt2ndUnit = txt1stUnit * 1
Else
If Volume(1) = "gallon (UK)" Then
  txt2ndUnit = txt1stUnit * 6.228835
Else
If Volume(1) = "gallon (US)" Then
  txt2ndUnit = txt1stUnit * 7.480519
End If
End If
End If
End If
End If
End If
End If


If Volume(0) = "gallon (UK)" Then
If Volume(1) = "litre (cu dm)" Then
 txt2ndUnit = txt1stUnit * 4.54609
Else
If Volume(1) = "cu metre" Then
  txt2ndUnit = txt1stUnit * 0.004546
Else
If Volume(1) = "cu inch" Then
  txt2ndUnit = txt1stUnit * 277.419433
Else
If Volume(1) = "cu foot" Then
  txt2ndUnit = txt1stUnit * 0.160544
Else
If Volume(1) = "gallon (UK)" Then
  txt2ndUnit = txt1stUnit * 1
Else
If Volume(1) = "gallon (US)" Then
  txt2ndUnit = txt1stUnit * 1.20095
End If
End If
End If
End If
End If
End If
End If


If Volume(0) = "gallon (US)" Then
If Volume(1) = "litre (cu dm)" Then
 txt2ndUnit = txt1stUnit * 3.785412
Else
If Volume(1) = "cu metre" Then
  txt2ndUnit = txt1stUnit * 0.003785
Else
If Volume(1) = "cu inch" Then
  txt2ndUnit = txt1stUnit * 231#
Else
If Volume(1) = "cu foot" Then
  txt2ndUnit = txt1stUnit * 0.133681
Else
If Volume(1) = "gallon (UK)" Then
  txt2ndUnit = txt1stUnit * 0.832674
Else
If Volume(1) = "gallon (US)" Then
  txt2ndUnit = txt1stUnit * 1
End If
End If
End If
End If
End If
End If
End If


End Sub
