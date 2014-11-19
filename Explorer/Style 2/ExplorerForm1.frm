VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "..:: My Explorer - Dhayalan Sivasuthan - www.dsiva.8m.com ::.."
   ClientHeight    =   9705
   ClientLeft      =   195
   ClientTop       =   825
   ClientWidth     =   9930
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ExplorerForm1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9705
   ScaleWidth      =   9930
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command28 
      Caption         =   "My Documents"
      Height          =   375
      Left            =   120
      TabIndex        =   87
      Top             =   9240
      Width           =   3255
   End
   Begin VB.CommandButton Command50 
      Caption         =   "Minimize All Windows"
      Height          =   375
      Left            =   6600
      TabIndex        =   29
      Top             =   9240
      Width           =   3255
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9015
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   15901
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Explorer"
      TabPicture(0)   =   "ExplorerForm1.frx":1CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label10"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label9"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label8"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Shape1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Shape2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Shape3"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblDriveLetter"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label13"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblDriveLabel"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label15"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblDriveType"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label17"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lblDriveTotSpace"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label19"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lblDriveFSpace"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label21"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Command49"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Command52"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Command12"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Command51"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Command5"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Command4"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Timer4"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Command33"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Command13"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Command6"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Timer2"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "tmrHeight"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "tmrWidth"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Text1"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Timer1"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "drvDrive"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "DirDirectories"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "filFile"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Frame3"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Frame1"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Command53"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).ControlCount=   41
      TabCaption(1)   =   "Quick Launch"
      TabPicture(1)   =   "ExplorerForm1.frx":1CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame2 
         Caption         =   "Other"
         ForeColor       =   &H000080FF&
         Height          =   3495
         Left            =   -74880
         TabIndex        =   71
         Top             =   5400
         Width           =   9375
         Begin VB.CommandButton Command45 
            BackColor       =   &H00C0C0C0&
            Caption         =   "About me"
            Height          =   495
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   86
            ToolTipText     =   "About the Author"
            Top             =   2760
            Width           =   8895
         End
         Begin VB.CommandButton Command26 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Turn off PC"
            Height          =   1095
            Left            =   7440
            Picture         =   "ExplorerForm1.frx":1D02
            Style           =   1  'Graphical
            TabIndex        =   81
            ToolTipText     =   "Turn your computer off"
            Top             =   1560
            Width           =   1695
         End
         Begin VB.CommandButton Command25 
            BackColor       =   &H00C0C0C0&
            Caption         =   "User Info"
            Height          =   1095
            Left            =   5640
            Picture         =   "ExplorerForm1.frx":25CC
            Style           =   1  'Graphical
            TabIndex        =   80
            ToolTipText     =   "Show some information about the current user loged on"
            Top             =   1560
            Width           =   1695
         End
         Begin VB.CommandButton Command22 
            BackColor       =   &H00C0C0C0&
            Caption         =   "IE Transparency"
            Height          =   1095
            Left            =   2040
            Picture         =   "ExplorerForm1.frx":2E96
            Style           =   1  'Graphical
            TabIndex        =   79
            ToolTipText     =   "Make Internet Explorer's window transparent so that user can see what is behind IE"
            Top             =   1560
            Width           =   1695
         End
         Begin VB.CommandButton Command21 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Windows startup"
            Height          =   1095
            Left            =   240
            Picture         =   "ExplorerForm1.frx":3760
            Style           =   1  'Graphical
            TabIndex        =   78
            ToolTipText     =   "Show the time at which Windows was logged on"
            Top             =   1560
            Width           =   1695
         End
         Begin VB.CommandButton Command17 
            BackColor       =   &H00C0C0C0&
            Caption         =   "CPU Speed"
            Height          =   1095
            Left            =   3840
            Picture         =   "ExplorerForm1.frx":402A
            Style           =   1  'Graphical
            TabIndex        =   77
            ToolTipText     =   "Show CPU Speed"
            Top             =   1560
            Width           =   1695
         End
         Begin VB.CommandButton Command14 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Start Button"
            Height          =   1095
            Left            =   7440
            Picture         =   "ExplorerForm1.frx":48F4
            Style           =   1  'Graphical
            TabIndex        =   76
            ToolTipText     =   "Show / Hide Taskbar"
            Top             =   360
            Width           =   1695
         End
         Begin VB.CommandButton Command20 
            BackColor       =   &H00C0C0C0&
            Caption         =   "CD Drive"
            Height          =   1095
            Left            =   5640
            Picture         =   "ExplorerForm1.frx":51BE
            Style           =   1  'Graphical
            TabIndex        =   75
            ToolTipText     =   "Change bouble click spped of mouse in milliseconds"
            Top             =   360
            Width           =   1695
         End
         Begin VB.CommandButton Command19 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Taskbar"
            Height          =   1095
            Left            =   3840
            Picture         =   "ExplorerForm1.frx":5A88
            Style           =   1  'Graphical
            TabIndex        =   74
            ToolTipText     =   "Show / Hide Taskbar"
            Top             =   360
            Width           =   1695
         End
         Begin VB.CommandButton Command18 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Clear Documents"
            Height          =   1095
            Left            =   2040
            Picture         =   "ExplorerForm1.frx":6352
            Style           =   1  'Graphical
            TabIndex        =   73
            ToolTipText     =   "Clears the recent document menu"
            Top             =   360
            Width           =   1695
         End
         Begin VB.CommandButton Command3 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Start Screen Saver now"
            Height          =   1095
            Left            =   240
            MaskColor       =   &H8000000F&
            Picture         =   "ExplorerForm1.frx":701C
            Style           =   1  'Graphical
            TabIndex        =   72
            ToolTipText     =   "Start the current screen saver immediately now"
            Top             =   360
            UseMaskColor    =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Control Panel"
         ForeColor       =   &H00FF0000&
         Height          =   2175
         Left            =   -74880
         TabIndex        =   61
         Top             =   480
         Width           =   9375
         Begin VB.CommandButton Command54 
            BackColor       =   &H00E0E0E0&
            Caption         =   "IE Properties"
            Height          =   855
            Left            =   7800
            Picture         =   "ExplorerForm1.frx":7CE6
            Style           =   1  'Graphical
            TabIndex        =   85
            ToolTipText     =   "Load Event Viewer"
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton Command38 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Folder Options"
            Height          =   855
            Left            =   7800
            Picture         =   "ExplorerForm1.frx":85B0
            Style           =   1  'Graphical
            TabIndex        =   84
            ToolTipText     =   "Load Command Pmompt"
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CommandButton Command10 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Control Panel"
            Height          =   1815
            Left            =   120
            Picture         =   "ExplorerForm1.frx":927A
            Style           =   1  'Graphical
            TabIndex        =   70
            ToolTipText     =   "Show Control Panel"
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton Command47 
            BackColor       =   &H00E0E0E0&
            Caption         =   "User Accounts"
            Height          =   855
            Left            =   6240
            Picture         =   "ExplorerForm1.frx":9B44
            Style           =   1  'Graphical
            TabIndex        =   69
            ToolTipText     =   "Load Command Pmompt"
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CommandButton Command37 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Volume Control"
            Height          =   855
            Left            =   4680
            Picture         =   "ExplorerForm1.frx":A80E
            Style           =   1  'Graphical
            TabIndex        =   68
            ToolTipText     =   "Load Volume Control"
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CommandButton Command42 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Taskbar"
            Height          =   855
            Left            =   3120
            Picture         =   "ExplorerForm1.frx":B0D8
            Style           =   1  'Graphical
            TabIndex        =   67
            ToolTipText     =   "Load Sound Recorder"
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CommandButton Command43 
            BackColor       =   &H00E0E0E0&
            Caption         =   "System Time"
            Height          =   855
            Left            =   1560
            Picture         =   "ExplorerForm1.frx":BDA2
            Style           =   1  'Graphical
            TabIndex        =   66
            ToolTipText     =   "Load Volume Control"
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CommandButton Command39 
            BackColor       =   &H00E0E0E0&
            Caption         =   "System"
            Height          =   855
            Left            =   6240
            Picture         =   "ExplorerForm1.frx":C66C
            Style           =   1  'Graphical
            TabIndex        =   65
            ToolTipText     =   "Load Event Viewer"
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton Command48 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Mouse"
            Height          =   855
            Left            =   4680
            Picture         =   "ExplorerForm1.frx":D336
            Style           =   1  'Graphical
            TabIndex        =   64
            ToolTipText     =   "Load Sound Recorder"
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton Command24 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Display"
            Height          =   855
            Left            =   3120
            Picture         =   "ExplorerForm1.frx":DC00
            Style           =   1  'Graphical
            TabIndex        =   63
            ToolTipText     =   "Show Display properties dialog"
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton Command46 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Add/Remove Programs"
            Height          =   855
            Left            =   1560
            Picture         =   "ExplorerForm1.frx":E4CA
            Style           =   1  'Graphical
            TabIndex        =   62
            ToolTipText     =   "Load Volume Control"
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Windows Applications and Tools"
         ForeColor       =   &H00008000&
         Height          =   2175
         Left            =   -74880
         TabIndex        =   49
         Top             =   3000
         Width           =   9375
         Begin VB.CommandButton Command40 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Help"
            Height          =   855
            Left            =   7800
            Picture         =   "ExplorerForm1.frx":F194
            Style           =   1  'Graphical
            TabIndex        =   82
            ToolTipText     =   "Load Paint"
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton Command36 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Sound Recorder"
            Height          =   855
            Left            =   120
            Picture         =   "ExplorerForm1.frx":FE5E
            Style           =   1  'Graphical
            TabIndex        =   60
            ToolTipText     =   "Load Sound Recorder"
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CommandButton Command31 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Command Prompt"
            Height          =   855
            Left            =   1680
            Picture         =   "ExplorerForm1.frx":10728
            Style           =   1  'Graphical
            TabIndex        =   59
            ToolTipText     =   "Load Command Pmompt"
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CommandButton Command16 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Run"
            Height          =   855
            Left            =   6360
            Picture         =   "ExplorerForm1.frx":10FF2
            Style           =   1  'Graphical
            TabIndex        =   58
            ToolTipText     =   "Show Run Dialog"
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton Command41 
            BackColor       =   &H00E0E0E0&
            Caption         =   "File Search"
            Height          =   855
            Left            =   4800
            Picture         =   "ExplorerForm1.frx":118BC
            Style           =   1  'Graphical
            TabIndex        =   57
            ToolTipText     =   "Load Notepad"
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton Command35 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Notepad"
            Height          =   855
            Left            =   3240
            Picture         =   "ExplorerForm1.frx":12586
            Style           =   1  'Graphical
            TabIndex        =   56
            ToolTipText     =   "Load Notepad"
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton Command34 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Paint"
            Height          =   855
            Left            =   1680
            Picture         =   "ExplorerForm1.frx":12E50
            Style           =   1  'Graphical
            TabIndex        =   55
            ToolTipText     =   "Load Paint"
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton Command27 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Calcularor"
            Height          =   855
            Left            =   120
            MaskColor       =   &H8000000F&
            Picture         =   "ExplorerForm1.frx":1371A
            Style           =   1  'Graphical
            TabIndex        =   54
            ToolTipText     =   "Load Calculator"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   1455
         End
         Begin VB.CommandButton Command32 
            BackColor       =   &H00E0E0E0&
            Caption         =   "DirectX Diag Tool"
            Height          =   855
            Left            =   6360
            Picture         =   "ExplorerForm1.frx":13FE4
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "Load DirectX Diagnostic Tool"
            Top             =   1200
            Width           =   1335
         End
         Begin VB.CommandButton Command30 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Disk Cleanup"
            Height          =   855
            Left            =   4800
            Picture         =   "ExplorerForm1.frx":148AE
            Style           =   1  'Graphical
            TabIndex        =   52
            ToolTipText     =   "Load Disk Cleanup"
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CommandButton Command29 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Character Map"
            Height          =   855
            Left            =   3240
            Picture         =   "ExplorerForm1.frx":15178
            Style           =   1  'Graphical
            TabIndex        =   51
            ToolTipText     =   "Load Character Map"
            Top             =   1200
            Width           =   1455
         End
         Begin VB.CommandButton Command15 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Recycle Bin"
            Height          =   855
            Left            =   7800
            Picture         =   "ExplorerForm1.frx":15A42
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Show Recycle Bin"
            Top             =   1200
            Width           =   1455
         End
      End
      Begin VB.CommandButton Command53 
         Caption         =   "Get Drive Info."
         Height          =   375
         Left            =   2880
         TabIndex        =   48
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Caption         =   "File Filter"
         Height          =   1095
         Left            =   8280
         TabIndex        =   34
         Top             =   7320
         Width           =   1215
         Begin VB.ComboBox Filter 
            Height          =   315
            ItemData        =   "ExplorerForm1.frx":1630C
            Left            =   120
            List            =   "ExplorerForm1.frx":1634F
            Style           =   2  'Dropdown List
            TabIndex        =   83
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Add ext"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   35
            Top             =   645
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Copy or Move Files"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1335
         Left            =   2400
         TabIndex        =   20
         Top             =   11400
         Width           =   10695
         Begin VB.TextBox Text3 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   465
            Left            =   1320
            MultiLine       =   -1  'True
            TabIndex        =   24
            Text            =   "ExplorerForm1.frx":163E4
            Top             =   225
            Width           =   9255
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   465
            Left            =   1320
            MultiLine       =   -1  'True
            TabIndex        =   23
            Text            =   "ExplorerForm1.frx":16427
            Top             =   795
            Width           =   9255
         End
         Begin VB.CommandButton Command7 
            BackColor       =   &H0080C0FF&
            Caption         =   "Set Source"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Sets the source file to be copied or moved"
            Top             =   300
            Width           =   1095
         End
         Begin VB.CommandButton Command11 
            BackColor       =   &H0080C0FF&
            Caption         =   "Set Target"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Sets the target file path to be copied or moved"
            Top             =   795
            Width           =   1095
         End
         Begin VB.Line Line3 
            BorderColor     =   &H000080FF&
            Visible         =   0   'False
            X1              =   0
            X2              =   8040
            Y1              =   735
            Y2              =   735
         End
      End
      Begin VB.FileListBox filFile 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   3150
         Left            =   240
         ReadOnly        =   0   'False
         System          =   -1  'True
         TabIndex        =   19
         Top             =   4800
         Width           =   7935
      End
      Begin VB.DirListBox DirDirectories 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   3015
         Left            =   4320
         TabIndex        =   18
         Top             =   1560
         Width           =   3855
      End
      Begin VB.DriveListBox drvDrive 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   240
         TabIndex        =   17
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   5730
         Top             =   7230
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000E&
         CausesValidation=   0   'False
         ForeColor       =   &H0000FF00&
         Height          =   675
         HideSelection   =   0   'False
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   600
         Width           =   8775
      End
      Begin VB.Timer tmrWidth 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   7080
         Top             =   7200
      End
      Begin VB.Timer tmrHeight 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   7440
         Top             =   7200
      End
      Begin VB.Timer Timer2 
         Interval        =   500
         Left            =   6240
         Top             =   7230
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Create Folder"
         Height          =   735
         Left            =   8280
         Picture         =   "ExplorerForm1.frx":16482
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Make a new folder"
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H00C0C0C0&
         Caption         =   "File Properties"
         Height          =   615
         Left            =   8280
         Picture         =   "ExplorerForm1.frx":16A0C
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Show file properties"
         Top             =   4800
         Width           =   1215
      End
      Begin VB.CommandButton Command33 
         BackColor       =   &H00C0C0C0&
         Caption         =   "<< Copy"
         Height          =   705
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   600
         Width           =   540
      End
      Begin VB.Timer Timer4 
         Interval        =   1000
         Left            =   3330
         Top             =   5190
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Copy File"
         Height          =   615
         Left            =   8280
         Picture         =   "ExplorerForm1.frx":16F96
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Copy the file,shown in source textbox, to the destination shown in target textbox"
         Top             =   5400
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Move File"
         Height          =   615
         Left            =   8280
         Picture         =   "ExplorerForm1.frx":17520
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Move the file,shown in source textbox, to the destination shown in target textbox"
         Top             =   6000
         Width           =   1215
      End
      Begin VB.CommandButton Command51 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Delete File"
         Height          =   615
         Left            =   8280
         Picture         =   "ExplorerForm1.frx":17AAA
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Move the file,shown in source textbox, to the destination shown in target textbox"
         Top             =   6600
         Width           =   1215
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Rename Folder"
         Height          =   735
         Left            =   8280
         Picture         =   "ExplorerForm1.frx":18034
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "rename the selected folder"
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton Command52 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Delete Folder"
         Height          =   735
         Left            =   8280
         Picture         =   "ExplorerForm1.frx":185BE
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "rename the selected folder"
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton Command49 
         BackColor       =   &H00808080&
         Caption         =   "Explore Folder"
         Height          =   735
         Left            =   8280
         Picture         =   "ExplorerForm1.frx":18B48
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label Label21 
         Caption         =   "Drive Used Space"
         Height          =   255
         Left            =   360
         TabIndex        =   47
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label lblDriveFSpace 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2040
         TabIndex        =   46
         Top             =   3600
         Width           =   1935
      End
      Begin VB.Label Label19 
         Caption         =   "Drive Free Space"
         Height          =   255
         Left            =   360
         TabIndex        =   45
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Label lblDriveTotSpace 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2040
         TabIndex        =   44
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label Label17 
         Caption         =   "Drive Total Space"
         Height          =   255
         Left            =   360
         TabIndex        =   43
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label lblDriveType 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2040
         TabIndex        =   42
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label15 
         Caption         =   "Drive Type"
         Height          =   255
         Left            =   360
         TabIndex        =   41
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label lblDriveLabel 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2040
         TabIndex        =   40
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label Label13 
         Caption         =   "Drive Label"
         Height          =   255
         Left            =   360
         TabIndex        =   39
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label lblDriveLetter 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2040
         TabIndex        =   38
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Drive Letter"
         Height          =   255
         Left            =   360
         TabIndex        =   37
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Shape Shape3 
         Height          =   3255
         Left            =   120
         Top             =   1440
         Width           =   4095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   6480
         TabIndex        =   36
         Top             =   8160
         Width           =   495
      End
      Begin VB.Shape Shape2 
         Height          =   3855
         Left            =   120
         Top             =   4680
         Width           =   9495
      End
      Begin VB.Shape Shape1 
         Height          =   3255
         Left            =   4200
         Top             =   1440
         Width           =   5415
      End
      Begin VB.Label Label2 
         Caption         =   "File Count:"
         Height          =   255
         Left            =   5640
         TabIndex        =   31
         Top             =   8160
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   1080
         TabIndex        =   28
         Top             =   8160
         Width           =   2295
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Modified Date:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   8040
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1097
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   4200
         TabIndex        =   26
         Top             =   8160
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Size:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3720
         TabIndex        =   25
         Top             =   8160
         Width           =   495
      End
   End
   Begin VB.CommandButton Command44 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Shutdown"
      Height          =   855
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Load Volume Control"
      Top             =   8160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command23 
      Caption         =   "IE version"
      Height          =   375
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   3
      Left            =   12360
      Top             =   3120
   End
   Begin VB.CommandButton Command9 
      Height          =   675
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Show / Hide extras"
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00D3D3D3&
      Caption         =   "Run"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10920
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Runs EXE files"
      Top             =   11400
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00D3D3D3&
      Caption         =   "Copy or move files > >"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   " Shows/Hides options for moving or copying files "
      Top             =   11400
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   7995
      Width           =   3135
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Label12"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      Top             =   11040
      Width           =   3855
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Recomended Resolution - 1024 by 768 pixels or better"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Left            =   6240
      TabIndex        =   6
      Top             =   9930
      Width           =   5415
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Dhayalan Sivasuthan"
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
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   9930
      Width           =   5415
   End
   Begin VB.Menu Explorer 
      Caption         =   "Explorer"
      Visible         =   0   'False
      Begin VB.Menu CORM 
         Caption         =   "Copy Or Move Files"
      End
      Begin VB.Menu Close 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu About 
      Caption         =   "About"
   End
   Begin VB.Menu Startbtn 
      Caption         =   "Startbtn"
      Visible         =   0   'False
      Begin VB.Menu Show 
         Caption         =   "Show"
      End
      Begin VB.Menu Hide 
         Caption         =   "Hide"
      End
   End
   Begin VB.Menu TaskbarSH 
      Caption         =   "Taskbar"
      Visible         =   0   'False
      Begin VB.Menu Showtask 
         Caption         =   "Show"
      End
      Begin VB.Menu Hidetask 
         Caption         =   "Hide"
      End
   End
   Begin VB.Menu MouseDblClick 
      Caption         =   "MouseDblClick"
      Visible         =   0   'False
      Begin VB.Menu mone 
         Caption         =   "100"
      End
      Begin VB.Menu mtwo 
         Caption         =   "200"
      End
      Begin VB.Menu mthree 
         Caption         =   "300"
      End
      Begin VB.Menu mfour 
         Caption         =   "400"
      End
      Begin VB.Menu mfive 
         Caption         =   "500"
      End
      Begin VB.Menu msix 
         Caption         =   "600"
      End
      Begin VB.Menu mseven 
         Caption         =   "700"
      End
      Begin VB.Menu meight 
         Caption         =   "800"
      End
   End
   Begin VB.Menu DisplayPro 
      Caption         =   "Display"
      Visible         =   0   'False
      Begin VB.Menu Scr 
         Caption         =   "Screen Saver"
      End
      Begin VB.Menu Apperance 
         Caption         =   "Apperance"
      End
      Begin VB.Menu Settings 
         Caption         =   "Settings"
      End
      Begin VB.Menu BG 
         Caption         =   "Background"
      End
   End
   Begin VB.Menu Applications 
      Caption         =   "Applications"
      Begin VB.Menu Calculator 
         Caption         =   "Calculator"
      End
      Begin VB.Menu Paint 
         Caption         =   "Paint"
      End
      Begin VB.Menu Notepad 
         Caption         =   "Notepad"
      End
      Begin VB.Menu WinUti 
         Caption         =   "Windows Utilities"
         Begin VB.Menu FSearch 
            Caption         =   "File Search"
         End
         Begin VB.Menu Run 
            Caption         =   "Run"
         End
         Begin VB.Menu Cmd 
            Caption         =   "Command Prompt"
         End
         Begin VB.Menu SRec 
            Caption         =   "Sound Recorder"
         End
         Begin VB.Menu CharMap 
            Caption         =   "Character Map"
         End
         Begin VB.Menu DCleanup 
            Caption         =   "Disk Cleanup"
         End
         Begin VB.Menu EeventVwr 
            Caption         =   "Event Viewer"
         End
         Begin VB.Menu DirectXDT 
            Caption         =   "DirectX Diag Tool"
         End
      End
      Begin VB.Menu CPanel 
         Caption         =   "Control Panel"
      End
      Begin VB.Menu CPanelItems 
         Caption         =   "Control Panel Items"
         Begin VB.Menu ARPrograms 
            Caption         =   "Add/Remove Programs"
         End
         Begin VB.Menu Display 
            Caption         =   "Display"
         End
         Begin VB.Menu Mouse 
            Caption         =   "Mouse"
         End
         Begin VB.Menu System 
            Caption         =   "System"
         End
         Begin VB.Menu SysTime 
            Caption         =   "System Time"
         End
         Begin VB.Menu Taskbar 
            Caption         =   "Taskbar"
         End
         Begin VB.Menu VControl 
            Caption         =   "Volume Control"
         End
         Begin VB.Menu UAccounts 
            Caption         =   "User Accounts"
         End
      End
      Begin VB.Menu WinHelp 
         Caption         =   "Windows Help"
      End
      Begin VB.Menu RBin 
         Caption         =   "Recycle Bin"
      End
      Begin VB.Menu Shutdown 
         Caption         =   "Shutdown Computer"
      End
   End
   Begin VB.Menu CD 
      Caption         =   "CD"
      Begin VB.Menu OpenCD 
         Caption         =   "Open"
      End
      Begin VB.Menu CloseCD 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Dim SH As New Shell  'reference to shell32.dll class
Dim ShBFF As Folder  'Shell Browse For Folder
'system startup
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
 

Private Sub Author_Click()
Dialog.Show
End Sub

Private Sub About_Click()
Dialog.Show
Me.Enabled = False
End Sub

Private Sub Apperance_Click()
ShowDisplayPropsDialog DP_Appearance
End Sub

Private Sub ARPrograms_Click()
Command46_Click
End Sub

Private Sub BG_Click()
ShowDisplayPropsDialog DP_Background
End Sub

Private Sub Calculator_Click()
Command27_Click
End Sub

Private Sub CharMap_Click()
Command29_Click
End Sub

Private Sub Close_Click()
Dim Ex As Integer
Ex = MsgBox("Do you really want to exit?", vbYesNo, "Siva's Explorer")
If Ex = vbYes Then
tmrHeight.Enabled = True
End If
End Sub

Private Sub CloseCD_Click()
Dim LRet As Long
    LRet = mciSendString("set CDAudio door closed", returnstring, 127, 0)

End Sub

Private Sub Cmd_Click()
Command31_Click
End Sub

Private Sub Command10_Click()
ShowControlPanel
End Sub

Private Sub Command11_Click()
Text2.Text = filFile.Path + "\" + filFile

End Sub

Private Sub Command12_Click()
On Error Resume Next
If filFile.Filename = "" Then
MsgBox "Select the file to be copied"
Exit Sub
End If

Dim RenDirName
RenDirName = InputBox("Give the new name for the folder you wanted to rename", "Rename Folder")
Name DirDirectories.Path As DirDirectories.Path & "\" & RenDirName  'for Rename Folder
End Sub

Private Sub Command13_Click()
If filFile.Filename = "" Then
MsgBox "Select the file"
Exit Sub
End If

ShowFileProperties filFile.Path + "\" + filFile.Filename, hWnd

End Sub

Private Sub Command14_Click()
PopupMenu Startbtn
End Sub

Private Sub Command15_Click()
ShowRecycleBin
End Sub

Private Sub Command16_Click()
Dim actionres As String
actionres = ShowRunDialog()
If actionres = False Then MsgBox "Show Run dialog failed"


End Sub

Private Sub Command17_Click()
MsgBox GetCPUSpeed() & " MHz"
End Sub

Private Sub Command18_Click()
ClearDocumentsMenu
MsgBox "Successfuly cleared"
End Sub

Private Sub Command19_Click()
PopupMenu Taskbar
End Sub

Private Sub Command20_Click()
PopupMenu CD
End Sub








Private Sub Command21_Click()
MsgBox GetSystemStartup()
End Sub

Private Sub Command22_Click()
If Timer3.Enabled = True Then
Timer3.Enabled = False
Else
Timer3.Enabled = True
End If
End Sub

Private Sub Command23_Click()
'MsgBox IEVersionShort
MsgBox IEVersionLong
End Sub


Private Sub Command24_Click()
PopupMenu DisplayPro
End Sub

Private Sub Command25_Click()
 'Set form autoredraw property to true to get this example
     'to work
     
     Dim myuser As Variant
     Dim I As Integer
     'used to test if any environment variables exist
     'sometimes they don't
     Dim sTemp As String
     
     I = 1
     myuser = Environ(I)
     sTemp = myuser
     MsgBox Environ(I)
     I = I + 1
     'Print all values on form
     Do While Len(myuser) > 0
          MsgBox Environ(I)
          I = I + 1
          myuser = Environ(I)
          sTemp = sTemp & myuser
     Loop
     If Len(sTemp) = 0 Then Form1.Print _
         "No environment variables exist"
     

End Sub

Private Sub Command26_Click()
Dim result As String
result = MsgBox("Do you really want to turn of you computer?", vbYesNo + vbQuestion, "Turn off ?")
If result = vbYes Then
    SH.ShutdownWindows
End If
End Sub



Private Sub Command27_Click()
Shell GetWinDir & "\" & "system32" & "\" & "calc.exe", vbNormalFocus
End Sub

Private Sub Command28_Click()
Shell GetWinDir & "\" & "explorer.exe", vbNormalFocus
End Sub

Private Sub Command29_Click()
Shell GetWinDir & "\" & "system32" & "\" & "charmap.exe", vbNormalFocus
End Sub

Private Sub Command30_Click()
Shell GetWinDir & "\" & "system32" & "\" & "cleanmgr.exe", vbNormalFocus
End Sub

Private Sub Command31_Click()
Shell GetWinDir & "\" & "system32" & "\" & "cmd.exe", vbNormalFocus
End Sub

Private Sub Command32_Click()
Shell GetWinDir & "\" & "system32" & "\" & "dxdiag.exe", vbNormalFocus
End Sub

Private Sub Command33_Click()
Clipboard.SetText Text1.Text
End Sub

Private Sub Command34_Click()
Shell GetWinDir & "\" & "system32" & "\" & "mspaint.exe", vbNormalFocus
End Sub

Private Sub Command35_Click()
Shell GetWinDir & "\" & "system32" & "\" & "notepad.exe", vbNormalFocus
End Sub

Private Sub Command36_Click()
Shell GetWinDir & "\" & "system32" & "\" & "sndrec32.exe", vbNormalFocus
End Sub

Private Sub Command37_Click()
Shell GetWinDir & "\" & "system32" & "\" & "sndvol32.exe", vbNormalFocus
End Sub

Private Sub Command38_Click()
Me.Enabled = False
Dialog.Show
End Sub

Private Sub Command39_Click()
  SH.ControlPanelItem "sysdm.cpl" 'System Properties

End Sub

Private Sub Command40_Click()
 SH.Help
End Sub

Private Sub Command41_Click()
SH.FindFiles
End Sub

Private Sub Command42_Click()
  SH.TrayProperties
End Sub

Private Sub Command43_Click()
  SH.ControlPanelItem "Timedate.cpl"
End Sub

Private Sub Command44_Click()

  If MsgBox("Are you sure you want to do this!?", _
     vbQuestion + vbYesNo + vbDefaultButton2, _
     "Confirm Action!") <> vbYes Then Exit Sub
     
    SH.ShutdownWindows

End Sub

Private Sub Command45_Click()
Command38_Click
End Sub

Private Sub Command46_Click()
SH.ControlPanelItem "appwiz.cpl"
End Sub

Private Sub Command47_Click()
SH.ControlPanelItem "nusrmgr.cpl"
End Sub

Private Sub Command48_Click()
SH.ControlPanelItem "main.cpl"
End Sub

Private Sub Command49_Click()
SH.Open DirDirectories.Path
End Sub

Private Sub Command50_Click()
  SH.MinimizeAll
End Sub



Private Sub Command51_Click()
Dim DeleteOk As String
DeleteOk = MsgBox("Do yo really want to delete the file " _
& filFile.Path & "\" & filFile.Filename & "?", vbYesNo + vbQuestion)
If DeleteOk = vbYes Then
ShellFileDelete filFile.Path & "\" & filFile.Filename, True
MsgBox "File Deleted"
End If
End Sub

Private Sub Command52_Click()
'If DirDirectories. = "" Then
'MsgBox "Select the file to be copied"
'Exit Sub
'End If
'File_Delete DirDirectories.Path
End Sub

Private Sub Command53_Click()
On Error Resume Next
lblDriveLetter.Caption = Left(drvDrive.Drive, 2)
lblDriveLabel.Caption = Mid(drvDrive.Drive, 3, (Len(drvDrive.Drive)))
lblDriveTotSpace.Caption = DriveMBSize(lblDriveLetter.Caption) & " MB"
lblDriveFSpace.Caption = DriveMBFree(lblDriveLetter.Caption) & " MB"
ProgressBar1.Max = Val(lblDriveTotSpace.Caption) / 100
ProgressBar1.Value = ((Val(lblDriveTotSpace.Caption) - Val(lblDriveFSpace.Caption)) / 100)
Select Case DriveType(lblDriveLetter.Caption)
Case 2: lblDriveType.Caption = "Removable"
Case 3: lblDriveType.Caption = "Fixed"
Case 5: lblDriveType.Caption = "CD ROM"
End Select
End Sub

Private Sub Command54_Click()
SH.ControlPanelItem "Inetcpl.cpl"
End Sub

Private Sub Command6_Click()
On Error Resume Next
Dim Fname  As String
Fname = InputBox("Give a name for the new folder", "Folder Name", "New Folder")
MkDir DirDirectories.Path & "\" & Fname  'for Make new Folder
DirDirectories.Refresh
End Sub

Private Sub CPanel_Click()
Command10_Click
End Sub

Private Sub DCleanup_Click()
Command30_Click
End Sub

Private Sub DirectXDT_Click()
Command32_Click
End Sub

Private Sub Display_Click()
SH.ControlPanelItem "desk.cpl"
End Sub

Private Sub EeventVwr_Click()
Command28_Click
End Sub

Private Sub filFile_DblClick()
If Right(filFile.Filename, 3) = "exe" Then
Shell filFile.Path + "\" + filFile.Filename, vbNormalFocus
Else
File_Open filFile.Path + "\" + filFile.Filename, "open"
End If
End Sub

Private Sub filter_Click()
filFile.Pattern = Filter.Text
End Sub

Private Sub Command1_Click()
On Error Resume Next
Dim Run As String
'If Right(Text1, 3) = "exe" Then
Run = Text1.Text
Shell Run, vbNormalFocus
'Else
'MsgBox "Error"
'End If
End Sub




Private Sub Command2_Click()
On Error Resume Next
Dim Ext As String
Ext = InputBox("Enter the Extention you want", "Siva's Explorer", "*.*")
Label3.Caption = Ext
filFile.Pattern = Label3.Caption
Filter.AddItem Label3.Caption
Filter.Text = Label3.Caption
End Sub




Private Sub Command3_Click()
  SendMessage Form1.hWnd, WM_SYSCOMMAND, _
      SC_SCREENSAVE, 0&
End Sub

Private Sub Command4_Click()
If filFile.Filename = "" Then
MsgBox "Select the file to be copied"
Exit Sub
End If

Dialog1.txtSrcFile.Text = Form1.Text1.Text
Me.Enabled = False
Dialog1.Show
End Sub

Private Sub Command5_Click()
If filFile.Filename = "" Then
MsgBox "Select the file to be moved"
End If
Dialog1.txtSrcFile.Text = Form1.Text1.Text
Me.Enabled = False
Dialog1.Show

End Sub



Private Sub Command7_Click()
Text3.Text = filFile.Path + "\" + filFile
End Sub



Private Sub Command8_Click()
If Me.Height = 4500 Then
Me.Height = 6105
Command8.Caption = "Copy or move files < <"
Else
Me.Height = 4500
Command8.Caption = "Copy or move files > >"
End If
End Sub

Private Sub Command9_Click()
Clipboard.SetText Text1.Text
End Sub





Private Sub CORM_Click()
Command8_Click
End Sub

Private Sub DirDirectories_Change()
filFile.Path = DirDirectories.List(DirDirectories.ListIndex)
Text1.Text = DirDirectories.List(DirDirectories.ListIndex)
Label4.Caption = ""
Label9.Caption = ""
Label1.Caption = filFile.ListCount

End Sub

Private Sub drvDrive_Change()
On Error GoTo DriveError
DirDirectories.Path = drvDrive.Drive
'If drvDrive.Drive = "d: [MY DOCUMENTS]" Then
'DirDirectories.Path = "\My Documents\My Music\Songs"
'End If
Exit Sub
DriveError:
MsgBox "Drive Inaccesible"
End Sub


Private Sub filFile_Click()
Text1 = filFile.Path + "\" + filFile
Label4.Caption = FileDateTime(Text1)
Label9.Caption = FileLen(Text1)
Label9.Caption = Format(Label9 / 1024, "0.000") + "KB"
Label1.Caption = filFile.ListCount
End Sub

Private Sub Form_Load()
'MsgBox "Welcome to Siva's Explorer", , "Siva's Explorer"
Text1.Text = filFile.Path
'control panel and etc


'Me.Height = 4500
End Sub




Private Sub Form_Unload(Cancel As Integer)
Dim Ex As Integer
Ex = MsgBox("Do you really want to exit?", vbYesNo, "Siva's Explorer")
If Ex = vbYes Then
Cancel = True
Dialog.Show
tmrHeight.Enabled = True
Else
Cancel = True
End If


End Sub












Private Sub FSearch_Click()
Command41_Click
End Sub

Private Sub Hidetask_Click()
HideTaskBar
End Sub

Private Sub Label20_Click()
End Sub

Private Sub Label6_Click()
Command38_Click
End Sub

Private Sub meight_Click()
SetDoubleClickTime (700)
End Sub

Private Sub mfive_Click()
SetDoubleClickTime (500)
End Sub

Private Sub mfour_Click()
SetDoubleClickTime (400)
End Sub

Private Sub mone_Click()
SetDoubleClickTime (100)
End Sub

Private Sub Mouse_Click()
Command48_Click
End Sub

Private Sub mseven_Click()
SetDoubleClickTime (700)
End Sub

Private Sub msix_Click()
SetDoubleClickTime (600)
End Sub

Private Sub mthree_Click()
SetDoubleClickTime (300)
End Sub

Private Sub mtwo_Click()
SetDoubleClickTime (200)
End Sub

Private Sub Notepad_Click()
Command35_Click
End Sub

Private Sub OpenCD_Click()
Dim LRet As Long
    LRet = mciSendString("set CDAudio door open", returnstring, 127, 0)

End Sub

Private Sub Paint_Click()
Command34_Click
End Sub

Private Sub RBin_Click()
Command15_Click
End Sub

Private Sub Run_Click()
Command16_Click
End Sub

Private Sub Scr_Click()
ShowDisplayPropsDialog DP_ScreenSaver
End Sub

Private Sub Settings_Click()
ShowDisplayPropsDialog DP_Settings
End Sub

Private Sub Showtask_Click()
ShowTaskBar
End Sub

Private Sub Shutdown_Click()
Command44_Click
End Sub

Private Sub SRec_Click()
Command36_Click
End Sub

Private Sub System_Click()
Command39_Click
End Sub

Private Sub SysTime_Click()
Command43_Click
End Sub

Private Sub Taskbar_Click()
Command42_Click
End Sub

Private Sub Timer3_Timer()
Dim lOldStyle As Long
    Dim bTrans As Byte ' The level of transparency (0 - 255)
    Dim LhWnd As Long
    
    LhWnd = FindWindow("IEFrame", vbNullString)
    bTrans = 200
        lOldStyle = GetWindowLong(LhWnd, GWL_EXSTYLE)
        SetWindowLong LhWnd, GWL_EXSTYLE, lOldStyle Or WS_EX_LAYERED
        SetLayeredWindowAttributes LhWnd, 0, bTrans, LWA_ALPHA

End Sub

Private Sub Timer4_Timer()
Label5.Caption = Now
End Sub

Private Sub tmrWidth_Timer()
If Form1.Width < 1770 Then
'tmrHeight.Enabled = True
End
Else
Form1.Width = Form1.Width - 60
End If
End Sub

Private Sub tmrHeight_Timer()
If Form1.Height < 600 Then
'tmrWidth.Enabled = True
End
Else
Me.WindowState = 0
Form1.Height = Form1.Height - 60
End If
End Sub

'************************************************************************************
'Other Functions
'-***********************************************************************************

Public Function ShowControlPanel() As Boolean
  On Error Resume Next
  Shell "rundll32 shell32,Control_RunDLL", vbNormalFocus
  ShowControlPanel = Err.Number = 0
End Function



Sub StartButton(blnValue As Boolean)
    Dim lngHandle As Long
    Dim lngStartButton As Long

    lngHandle = FindWindow("Shell_TrayWnd", "")
    lngStartButton = FindWindowEx(lngHandle, 0, "Button", _
    vbNullString)

    If blnValue Then
        ShowWindow lngStartButton, 5
    Else
        ShowWindow lngStartButton, 0
    End If

End Sub

Public Function GetWinDir() As String
    Dim sRet As String, lngLen As Long
    sRet = String(255, 0)
    lngLen = GetWindowsDirectory(sRet, 255)
    If lngLen = 0 Then Err.Raise Err.LastDllError
    GetWinDir = Left$(sRet, lngLen)
End Function

Private Sub Show_Click()
Dim a As Boolean
a = True
Call StartButton(a)
End Sub

Private Sub Hide_click()
Dim a As Boolean
a = False
Call StartButton(a)
'Me.Hide
End Sub


'show run dialog
Public Function ShowRunDialog() As Boolean

Dim oShellApp As Object
On Error Resume Next
Set oShellApp = CreateObject("Shell.Application")
oShellApp.FileRun
ShowRunDialog = Err.Number = 0
End Function

'recyclebin
Public Function ShowRecycleBin() As Boolean
      Dim LRet As Long
     'if using from a form, you can use me.hwnd instead of 0&
     'for the first argument
       LRet = ShellExecute(0&, "Open", "explorer.exe", _
       "/root,::{645FF040-5081-101B-9F08-00AA002F954E}", 0&, _
        SW_SHOWNORMAL)
        ShowRecycleBin = LRet > 32
End Function

  Private Function GetCPUSpeed() As Long
 
    Dim hKey As Long
   Dim CPUSpeed As Long

   Call RegOpenKey(HKEY_LOCAL_MACHINE, sCPURegKey, hKey)
                    
   Call RegQueryValueEx(hKey, "~MHz", 0, 0, CPUSpeed, 4)
   Call RegCloseKey(hKey)
    
   GetCPUSpeed = CPUSpeed
   
  End Function

'time win running
'Public Function WindowsRunTime() As Long
   ' WindowsRunTime = GetTickCount()
'End Function

Public Function ClearDocumentsMenu() As Boolean
    'Returns true if successful, false otherwise
    SHAddToRecentDocs 2, vbNullString
    ClearDocumentsMenu = Err.LastDllError = 0
End Function

'task bar hide or show
Public Function HideTaskBar() As Boolean
Dim LRet As Long
    LRet = FindWindow("Shell_traywnd", "")
    If LRet > 0 Then
        LRet = SetWindowPos(LRet, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
        HideTaskBar = LRet > 0
    End If
End Function

Public Function ShowTaskBar() As Boolean
Dim LRet As Long
LRet = FindWindow("Shell_traywnd", "")
If LRet > 0 Then
    LRet = SetWindowPos(LRet, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
    ShowTaskBar = LRet > 0
End If
End Function

'open with dialog
Public Sub DisplayOpenWith(strFile As String)

'***PURPOSE: DISPLAY OPEN WITH DIALOG:
'   PASS IT A FILE NAME
'   e.g., DisplayOpenWith "C:\FileWithNoDefaultApplication.bvq"
'**************************************

    On Error Resume Next
    Shell "rundll32.exe shell32.dll, OpenAs_RunDLL " & strFile

End Sub

'sys startup
Public Function GetSystemStartup() As Date

   Dim dTicks     As Double

      'Store the number of days the systems has been running
   dTicks = GetTickCount / 1000 / 60 / 60 / 24

   GetSystemStartup = Now() - dTicks
End Function

'ie version
Public Function IEVersionShort() As Long
    Dim udtVersionInfo As DllVersionInfo
    udtVersionInfo.cbSize = Len(udtVersionInfo)
    Call DllGetVersion(udtVersionInfo)
    IEVersionShort = udtVersionInfo.dwMajorVersion
End Function


Public Function IEVersionLong() As String
    Dim udtVersionInfo As DllVersionInfo
    udtVersionInfo.cbSize = Len(udtVersionInfo)
    Call DllGetVersion(udtVersionInfo)
    IEVersionLong = "Internet Explorer " & _
    udtVersionInfo.dwMajorVersion & "." & _
    udtVersionInfo.dwMinorVersion & "." & _
    udtVersionInfo.dwBuildNumber
End Function



'display properties
Public Function ShowDisplayPropsDialog(SubDialog As _
  DISPLAY_PROPERTIES) As Boolean
   
'Purpose: Shows Display Settings Dialog
'Parameter: Which Tab to show. Refer to Enumeration for details
'Returns: True if successful, false otherwis
'Example: ShowDisplayPropsDialog DP_Settings:
'         This displays settings dialog

If SubDialog < DP_Background Or SubDialog > DP_Settings Then
    'invalid parameter, default to settings
    SubDialog = DP_Settings
End If

Shell "rundll32.exe shell32.dll,Control_RunDLL desk.cpl,," & _
  SubDialog

ShowDisplayPropsDialog = Err.Number = 0 And Err.LastDllError = 0
 
End Function







Private Sub UAccounts_Click()
Command47_Click
End Sub

Private Sub VControl_Click()
Command37_Click
End Sub

Private Sub WinHelp_Click()
Command40_Click
End Sub
