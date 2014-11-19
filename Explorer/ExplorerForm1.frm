VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "..:: My Explorer - Dhayalan Sivasuthan - www.dsiva.8m.com ::.."
   ClientHeight    =   10065
   ClientLeft      =   195
   ClientTop       =   825
   ClientWidth     =   11850
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
   ScaleHeight     =   10065
   ScaleWidth      =   11850
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6735
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   11880
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "ExplorerForm1.frx":1CCA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label11"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label10"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label9"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label8"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command49"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Timer4"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command33"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command4"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command5"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Command13"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Command12"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Command6"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Timer2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "tmrHeight"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "tmrWidth"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Timer1"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "drvDrive"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "DirDirectories"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "filFile"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Frame1"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Frame3"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).ControlCount=   24
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "ExplorerForm1.frx":1CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Command18"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Command19"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Command20"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Command14"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Command17"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Command21"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Command22"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Command25"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Command26"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Command38"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "ExplorerForm1.frx":1D02
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   3015
         Left            =   0
         TabIndex        =   64
         Top             =   480
         Width           =   7335
         Begin VB.CommandButton Command10 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Control Panel"
            Height          =   855
            Left            =   120
            Picture         =   "ExplorerForm1.frx":1D1E
            Style           =   1  'Graphical
            TabIndex        =   73
            ToolTipText     =   "Show Control Panel"
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton Command46 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Add/Remove Programs"
            Height          =   855
            Left            =   1440
            Picture         =   "ExplorerForm1.frx":25E8
            Style           =   1  'Graphical
            TabIndex        =   72
            ToolTipText     =   "Load Volume Control"
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton Command24 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Display"
            Height          =   855
            Left            =   2640
            Picture         =   "ExplorerForm1.frx":2EB2
            Style           =   1  'Graphical
            TabIndex        =   71
            ToolTipText     =   "Show Display properties dialog"
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton Command48 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Mouse"
            Height          =   855
            Left            =   3960
            Picture         =   "ExplorerForm1.frx":377C
            Style           =   1  'Graphical
            TabIndex        =   70
            ToolTipText     =   "Load Sound Recorder"
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton Command39 
            BackColor       =   &H00E0E0E0&
            Caption         =   "System"
            Height          =   855
            Left            =   5280
            Picture         =   "ExplorerForm1.frx":4046
            Style           =   1  'Graphical
            TabIndex        =   69
            ToolTipText     =   "Load Event Viewer"
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton Command43 
            BackColor       =   &H00E0E0E0&
            Caption         =   "System Time"
            Height          =   855
            Left            =   240
            Picture         =   "ExplorerForm1.frx":4D10
            Style           =   1  'Graphical
            TabIndex        =   68
            ToolTipText     =   "Load Volume Control"
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton Command42 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Taskbar"
            Height          =   855
            Left            =   1800
            Picture         =   "ExplorerForm1.frx":55DA
            Style           =   1  'Graphical
            TabIndex        =   67
            ToolTipText     =   "Load Sound Recorder"
            Top             =   1920
            Width           =   1215
         End
         Begin VB.CommandButton Command37 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Volume Control"
            Height          =   855
            Left            =   3240
            Picture         =   "ExplorerForm1.frx":62A4
            Style           =   1  'Graphical
            TabIndex        =   66
            ToolTipText     =   "Load Volume Control"
            Top             =   1920
            Width           =   1215
         End
         Begin VB.CommandButton Command47 
            BackColor       =   &H00E0E0E0&
            Caption         =   "User Accounts"
            Height          =   855
            Left            =   4680
            Picture         =   "ExplorerForm1.frx":6B6E
            Style           =   1  'Graphical
            TabIndex        =   65
            ToolTipText     =   "Load Command Pmompt"
            Top             =   1920
            Width           =   1215
         End
      End
      Begin VB.CommandButton Command38 
         BackColor       =   &H00C0C0C0&
         Caption         =   "About me"
         Height          =   975
         Left            =   -74880
         Picture         =   "ExplorerForm1.frx":7838
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "About the Author"
         Top             =   5280
         Width           =   3135
      End
      Begin VB.CommandButton Command26 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Turn off PC"
         Height          =   975
         Left            =   -73320
         Picture         =   "ExplorerForm1.frx":9502
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Turn your computer off"
         Top             =   4320
         Width           =   1575
      End
      Begin VB.CommandButton Command25 
         BackColor       =   &H00C0C0C0&
         Caption         =   "User Info"
         Height          =   975
         Left            =   -74880
         Picture         =   "ExplorerForm1.frx":9DCC
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Show some information about the current user loged on"
         Top             =   4320
         Width           =   1575
      End
      Begin VB.CommandButton Command22 
         BackColor       =   &H00C0C0C0&
         Caption         =   "IE Transparency"
         Height          =   975
         Left            =   -73320
         Picture         =   "ExplorerForm1.frx":A696
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Make Internet Explorer's window transparent so that user can see what is behind IE"
         Top             =   3360
         Width           =   1575
      End
      Begin VB.CommandButton Command21 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Windows startup"
         Height          =   975
         Left            =   -73320
         Picture         =   "ExplorerForm1.frx":AF60
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Show the time at which Windows was logged on"
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton Command17 
         BackColor       =   &H00C0C0C0&
         Caption         =   "CPU Speed"
         Height          =   975
         Left            =   -74880
         Picture         =   "ExplorerForm1.frx":B82A
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Show CPU Speed"
         Top             =   3360
         Width           =   1575
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Start Button"
         Height          =   975
         Left            =   -74880
         Picture         =   "ExplorerForm1.frx":C0F4
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Show / Hide Taskbar"
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton Command20 
         BackColor       =   &H00C0C0C0&
         Caption         =   "DoubleClickSpeed"
         Height          =   975
         Left            =   -73320
         Picture         =   "ExplorerForm1.frx":C9BE
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Change bouble click spped of mouse in milliseconds"
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton Command19 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Taskbar"
         Height          =   975
         Left            =   -74880
         Picture         =   "ExplorerForm1.frx":D288
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Show / Hide Taskbar"
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton Command18 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Clear Documents"
         Height          =   975
         Left            =   -73320
         Picture         =   "ExplorerForm1.frx":DB52
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Clears the recent document menu"
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Start Screen Saver now"
         Height          =   975
         Left            =   -74880
         MaskColor       =   &H8000000F&
         Picture         =   "ExplorerForm1.frx":E81C
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Start the current screen saver immediately now"
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   1575
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00000000&
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
         ForeColor       =   &H0000FF00&
         Height          =   1335
         Left            =   -74880
         TabIndex        =   42
         Top             =   5880
         Width           =   8055
         Begin VB.TextBox Text3 
            BackColor       =   &H00000000&
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
            Height          =   465
            Left            =   1320
            MultiLine       =   -1  'True
            TabIndex        =   46
            Text            =   "ExplorerForm1.frx":F4E6
            Top             =   225
            Width           =   6615
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00000000&
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
            Height          =   465
            Left            =   1320
            MultiLine       =   -1  'True
            TabIndex        =   45
            Text            =   "ExplorerForm1.frx":F529
            Top             =   795
            Width           =   6615
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
            TabIndex        =   44
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
            TabIndex        =   43
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
      Begin VB.Frame Frame1 
         BackColor       =   &H00000000&
         Caption         =   "Filters files in the filelist to given ext."
         ForeColor       =   &H0000FFFF&
         Height          =   630
         Left            =   -69720
         TabIndex        =   38
         Top             =   5040
         Width           =   2895
         Begin VB.CommandButton Command2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Add extension"
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
            Left            =   1320
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   240
            Width           =   1455
         End
         Begin VB.ComboBox Filter 
            BackColor       =   &H00000000&
            ForeColor       =   &H0000FF00&
            Height          =   315
            ItemData        =   "ExplorerForm1.frx":F586
            Left            =   120
            List            =   "ExplorerForm1.frx":F5A5
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Label1"
            Height          =   495
            Left            =   0
            TabIndex        =   41
            Top             =   1560
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin VB.FileListBox filFile 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   2625
         Left            =   -70800
         ReadOnly        =   0   'False
         System          =   -1  'True
         TabIndex        =   37
         Top             =   2295
         Width           =   3975
      End
      Begin VB.DirListBox DirDirectories 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   2115
         Left            =   -74880
         TabIndex        =   36
         Top             =   2655
         Width           =   3975
      End
      Begin VB.DriveListBox drvDrive 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   315
         Left            =   -74880
         TabIndex        =   35
         Top             =   2295
         Width           =   3975
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   -72000
         Top             =   6240
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         CausesValidation=   0   'False
         ForeColor       =   &H0000FFFF&
         Height          =   675
         HideSelection   =   0   'False
         Left            =   -74160
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1560
         Width           =   7335
      End
      Begin VB.Timer tmrWidth 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   -70200
         Top             =   6120
      End
      Begin VB.Timer tmrHeight 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   -69840
         Top             =   6120
      End
      Begin VB.Timer Timer2 
         Interval        =   500
         Left            =   -71400
         Top             =   6240
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "New Folder"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   -74880
         Picture         =   "ExplorerForm1.frx":F5E6
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Make a new folder"
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Rename Folder"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   -73440
         Picture         =   "ExplorerForm1.frx":FEB0
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "rename the selected folder"
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H00C0C0C0&
         Caption         =   "File Properties"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   -71760
         Picture         =   "ExplorerForm1.frx":1077A
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Show file properties"
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Move File"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   -70080
         Picture         =   "ExplorerForm1.frx":11044
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Move the file,shown in source textbox, to the destination shown in target textbox"
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Copy File"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   -68400
         Picture         =   "ExplorerForm1.frx":1190E
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Copy the file,shown in source textbox, to the destination shown in target textbox"
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton Command33 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Copy>"
         Height          =   345
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1875
         Width           =   615
      End
      Begin VB.Timer Timer4 
         Interval        =   1000
         Left            =   -67800
         Top             =   2280
      End
      Begin VB.CommandButton Command49 
         BackColor       =   &H00808080&
         Caption         =   "Explore with Windows Explorer"
         Height          =   300
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   4800
         Width           =   3990
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
         Left            =   -74040
         TabIndex        =   52
         Top             =   5400
         Width           =   1935
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
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   -74880
         TabIndex        =   51
         Top             =   5280
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
         Left            =   -71205
         TabIndex        =   50
         Top             =   5400
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Size"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   -71640
         TabIndex        =   49
         Top             =   5400
         Width           =   495
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "KB"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   -70080
         TabIndex        =   48
         Top             =   5400
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000000&
         Caption         =   "Address:"
         ForeColor       =   &H0000FF00&
         Height          =   495
         Left            =   -74880
         TabIndex        =   47
         Top             =   1560
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command44 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Shutdown"
      Height          =   855
      Left            =   10200
      Picture         =   "ExplorerForm1.frx":121D8
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Load Volume Control"
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Recycle Bin"
      Height          =   855
      Left            =   8760
      Picture         =   "ExplorerForm1.frx":12AA2
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Show Recycle Bin"
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton Command40 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Help"
      Height          =   855
      Left            =   5880
      Picture         =   "ExplorerForm1.frx":1336C
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Load Paint"
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton Command32 
      BackColor       =   &H00E0E0E0&
      Caption         =   "DirectX Diag Tool"
      Height          =   855
      Left            =   4440
      Picture         =   "ExplorerForm1.frx":14036
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Load DirectX Diagnostic Tool"
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton Command28 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Event Viewer"
      Height          =   855
      Left            =   3000
      Picture         =   "ExplorerForm1.frx":14900
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Load Event Viewer"
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton Command30 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Disk Cleanup"
      Height          =   855
      Left            =   1560
      Picture         =   "ExplorerForm1.frx":151CA
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Load Disk Cleanup"
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton Command29 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Character Map"
      Height          =   855
      Left            =   120
      Picture         =   "ExplorerForm1.frx":15A94
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Load Character Map"
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton Command45 
      BackColor       =   &H00C0C0C0&
      Caption         =   "About me"
      Height          =   855
      Left            =   10200
      Picture         =   "ExplorerForm1.frx":1635E
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "About the Author"
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CommandButton Command36 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Sound Recorder"
      Height          =   855
      Left            =   8760
      Picture         =   "ExplorerForm1.frx":18028
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Load Sound Recorder"
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CommandButton Command31 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Command Prompt"
      Height          =   855
      Left            =   7320
      Picture         =   "ExplorerForm1.frx":188F2
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Load Command Pmompt"
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Run"
      Height          =   855
      Left            =   5880
      Picture         =   "ExplorerForm1.frx":191BC
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Show Run Dialog"
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CommandButton Command41 
      BackColor       =   &H00E0E0E0&
      Caption         =   "File Search"
      Height          =   855
      Left            =   4440
      Picture         =   "ExplorerForm1.frx":19A86
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Load Notepad"
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CommandButton Command35 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Notepad"
      Height          =   855
      Left            =   3000
      Picture         =   "ExplorerForm1.frx":1A750
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Load Notepad"
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CommandButton Command34 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Paint"
      Height          =   855
      Left            =   1560
      Picture         =   "ExplorerForm1.frx":1B01A
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Load Paint"
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CommandButton Command27 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Calcularor"
      Height          =   855
      Left            =   120
      MaskColor       =   &H8000000F&
      Picture         =   "ExplorerForm1.frx":1B8E4
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Load Calculator"
      Top             =   7200
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Some Does"
      ForeColor       =   &H0000FFFF&
      Height          =   6975
      Left            =   8400
      TabIndex        =   4
      Top             =   120
      Width           =   3375
      Begin VB.CommandButton Command50 
         Caption         =   "Minimize All Windows"
         Height          =   330
         Left            =   120
         TabIndex        =   25
         Top             =   6090
         Width           =   3135
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
         TabIndex        =   20
         Top             =   6525
         Width           =   3135
      End
   End
   Begin VB.CommandButton Command23 
      Caption         =   "IE version"
      Height          =   375
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   3
      Left            =   12360
      Top             =   3000
   End
   Begin VB.CommandButton Command9 
      Height          =   675
      Left            =   12480
      Picture         =   "ExplorerForm1.frx":1C1AE
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Show / Hide extras"
      Top             =   480
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
      Top             =   11280
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
      Top             =   11280
      Visible         =   0   'False
      Width           =   2775
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
      TabIndex        =   24
      Top             =   10920
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
      TabIndex        =   23
      Top             =   9810
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
      TabIndex        =   22
      Top             =   9810
      Width           =   5415
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FFFF&
      X1              =   1920
      X2              =   11640
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      X1              =   8280
      X2              =   8280
      Y1              =   120
      Y2              =   6960
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Windows Components"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   6960
      Width           =   2295
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
Dim RenDirName
RenDirName = InputBox("Give the new name for the folder you wanted to rename", "Rename Folder")
Name DirDirectories.Path As DirDirectories.Path & "\" & RenDirName  'for Rename Folder
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
PopupMenu MouseDblClick
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
Shell GetWinDir & "\" & "system32" & "\" & "eventvwr.exe", vbNormalFocus
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
  SH.SetTime
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
If Text2.Text = "" Or Text3.Text = "" Then
MsgBox "Set source and target paths"
Exit Sub
End If

Dim Source As String
Dim Target As String
Dim a As String
Dim S As String
Dim T As String

S = Text3.Text
T = Text2.Text

Source = S
Target = T + filFile

'Copy File
a = CopyFile(Trim$(Source), Trim(Target), False)
If a Then
        MsgBox "File copied!"
Else
        MsgBox "Error. File not moved!"
End If




End Sub

Private Sub Command5_Click()
If Text2.Text = "" Or Text3.Text = "" Then
MsgBox "Set source and target paths"
Exit Sub
End If
Dim Source As String
Dim Target As String
Dim a As String
Dim S As String
Dim T As String

S = Text3.Text
T = Text2.Text

Source = S
Target = T + filFile


'Move File
a = MoveFile(Trim$(Source), Trim(Target))
If a Then
        MsgBox "File moved!"
Else
        MsgBox "Error. File not moved!"
End If
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
Label9.Caption = Format(Label9 / 1024, "0.000")

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
      Dim lRet As Long
     'if using from a form, you can use me.hwnd instead of 0&
     'for the first argument
       lRet = ShellExecute(0&, "Open", "explorer.exe", _
       "/root,::{645FF040-5081-101B-9F08-00AA002F954E}", 0&, _
        SW_SHOWNORMAL)
        ShowRecycleBin = lRet > 32
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
Dim lRet As Long
    lRet = FindWindow("Shell_traywnd", "")
    If lRet > 0 Then
        lRet = SetWindowPos(lRet, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
        HideTaskBar = lRet > 0
    End If
End Function

Public Function ShowTaskBar() As Boolean
Dim lRet As Long
lRet = FindWindow("Shell_traywnd", "")
If lRet > 0 Then
    lRet = SetWindowPos(lRet, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
    ShowTaskBar = lRet > 0
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
