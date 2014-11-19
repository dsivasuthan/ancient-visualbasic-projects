VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL3N.OCX"
Begin VB.Form Form1 
   Caption         =   "My Text Editor - Dhayalan Sivasuthan"
   ClientHeight    =   6600
   ClientLeft      =   900
   ClientTop       =   1110
   ClientWidth     =   10110
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "TextEditorForm1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   10110
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   27
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Creates a new document"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Open a new document"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Saves the current opened document"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Bolds the current selection"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Italiise selected text"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Underlines selected text"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Makes the selected text unformated"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Change Color"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Format"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Cut"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Copy"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Paste"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Increases the size of the selected text by one point"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Decreases the size of the selected text by one point"
            Object.Tag             =   ""
            ImageIndex      =   14
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Finds a given string"
            Object.Tag             =   ""
            ImageIndex      =   21
         EndProperty
         BeginProperty Button20 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button21 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Insert Date"
            Object.Tag             =   ""
            ImageIndex      =   22
         EndProperty
         BeginProperty Button22 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button23 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Subscript"
            Object.Tag             =   ""
            ImageIndex      =   40
         EndProperty
         BeginProperty Button24 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Superscript"
            Object.Tag             =   ""
            ImageIndex      =   41
         EndProperty
         BeginProperty Button25 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button26 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Simlify"
            Object.Tag             =   ""
            ImageIndex      =   43
         EndProperty
         BeginProperty Button27 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Capitalize"
            Object.Tag             =   ""
            ImageIndex      =   42
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.ComboBox cmbFSize 
      Height          =   315
      Left            =   4320
      TabIndex        =   11
      Top             =   525
      Width           =   1095
   End
   Begin VB.ComboBox cmbFont 
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   525
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Caption         =   "HTML tags"
      Height          =   5415
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1215
      Begin VB.CommandButton Command6 
         Caption         =   "<font>"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "<title>"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "<body>"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "<html>"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   375
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3720
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Text Files|*.txt|RichTextFile|*.rtf"
      Flags           =   7
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   5535
      Left            =   1440
      TabIndex        =   3
      Top             =   960
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   9763
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"TextEditorForm1.frx":0CCA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   8520
      Top             =   5160
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Open"
      Height          =   375
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   735
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   6120
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   43
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":0D45
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":0F1F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":10F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":12D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":14AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":1687
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":1861
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":1A3B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":1C15
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":1DEF
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":1FC9
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":21A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":237D
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":2557
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":2731
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":290B
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":2AE5
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":2CBF
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":2E99
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":3073
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":324D
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":3427
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":3601
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":37DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":39B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":3B8F
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":3D69
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":3F43
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":411D
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":42F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":4611
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":47EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":49C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":4B9F
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":4D79
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":4F53
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":512D
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":5307
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":54E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":56BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":5895
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":5A6F
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "TextEditorForm1.frx":5C49
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape1 
      Height          =   375
      Left            =   8640
      Top             =   600
      Visible         =   0   'False
      Width           =   8535
   End
   Begin VB.Label Label1 
      Caption         =   "Click Open to open a html, text or a word document file using Siva's Text Editor"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   1440
      Width           =   8415
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu New 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu Open 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu Save 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu SaveAs 
         Caption         =   "Save As"
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "Edit"
      Begin VB.Menu Format 
         Caption         =   "Format"
      End
      Begin VB.Menu Space 
         Caption         =   "-"
      End
      Begin VB.Menu Bold 
         Caption         =   "Bold"
         Shortcut        =   ^B
      End
      Begin VB.Menu Italic 
         Caption         =   "Italic"
         Shortcut        =   ^I
      End
      Begin VB.Menu Underline 
         Caption         =   "Underline"
         Shortcut        =   ^U
      End
      Begin VB.Menu Regular 
         Caption         =   "Regular"
      End
      Begin VB.Menu Space4 
         Caption         =   "-"
      End
      Begin VB.Menu Fontcolor 
         Caption         =   "Font Color"
      End
      Begin VB.Menu Space5 
         Caption         =   "-"
      End
      Begin VB.Menu InsertDate 
         Caption         =   "Insert Date"
         Begin VB.Menu InsertDateDate 
            Caption         =   "Date"
         End
         Begin VB.Menu DateSampleOne 
            Caption         =   "27 February 2008"
            Visible         =   0   'False
         End
         Begin VB.Menu DateSampleTwo 
            Caption         =   "27.02.2008"
            Visible         =   0   'False
         End
         Begin VB.Menu DateSampleThree 
            Caption         =   "27-02-2008"
            Visible         =   0   'False
         End
         Begin VB.Menu InsertDateTime 
            Caption         =   "Time"
         End
         Begin VB.Menu InsertDateAndTime 
            Caption         =   "Date and Time"
         End
      End
      Begin VB.Menu SelectAll 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu ClearSelection 
         Caption         =   "Clear selection"
      End
      Begin VB.Menu space3 
         Caption         =   "-"
      End
      Begin VB.Menu Find 
         Caption         =   "Find"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu About 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SaveLocation As String
Dim Saved As Boolean
Dim SaveType As String



Private Sub About_Click()
Me.Enabled = False
Dialog.Show

End Sub

Private Sub Bold_Click()
RichTextBox1.SelBold = Not RichTextBox1.SelBold
End Sub

Private Sub ClearSelection_Click()
RichTextBox1.SelText = ""
End Sub



Private Sub cmbFont_Change()
RichTextBox1.SelFontName = cmbFont.Text
End Sub

Private Sub cmbFont_Scroll()
RichTextBox1.SelFontName = cmbFont.Text
End Sub

Private Sub cmbFSize_Change()
RichTextBox1.SelFontSize = cmbFSize.Text
End Sub

Private Sub cmbFSize_Scroll()
RichTextBox1.SelFontSize = cmbFSize.Text
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackColor = &H80FF&
End Sub

Private Sub Command2_Click()
On Error Resume Next
CommonDialog1.ShowOpen
RichTextBox1.LoadFile CommonDialog1.FileName
SaveLocation = CommonDialog1.FileName
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.BackColor = &H80FF&
End Sub



Private Sub DateAndTime_Click()
RichTextBox1.SelText = Now
End Sub

Private Sub Command3_Click()
Dim posi1 As String
posi1 = RichTextBox1.SelStart
RichTextBox1.SetFocus
RichTextBox1.SelText = "<html>" & vbCrLf & vbCrLf & "</html>"
RichTextBox1.SetFocus
RichTextBox1.SelStart = posi1 + 9
End Sub

Private Sub Command4_Click()
Dim posi1 As String
posi1 = RichTextBox1.SelStart
RichTextBox1.SetFocus
RichTextBox1.SelText = "<body>" & vbCrLf & vbCrLf & "</body>"
RichTextBox1.SetFocus
RichTextBox1.SelStart = posi1 + 9

End Sub

Private Sub Command5_Click()
Dim posi1 As String
posi1 = RichTextBox1.SelStart
RichTextBox1.SetFocus
RichTextBox1.SelText = "<title>" & vbCrLf & vbCrLf & "</title>"
RichTextBox1.SetFocus
RichTextBox1.SelStart = posi1 + 9

End Sub

Private Sub Command6_Click()
Dim posi1 As String
posi1 = RichTextBox1.SelStart
RichTextBox1.SetFocus
RichTextBox1.SelText = "<font>" & vbCrLf & vbCrLf & "</font>"
RichTextBox1.SetFocus
RichTextBox1.SelStart = posi1 + 9

End Sub

Private Sub DateSampleOne_Click()
'Dim CrtDate As Date
'CrtDate = Now
'CrtDate = Format$(CrtDate, "##")
'Format$(CrtDate, "dd mmmm yyyy")
RichTextBox1.SelText = CrtDate
End Sub

Private Sub Find_Click()
Form2.Show
End Sub

Private Sub Fontcolor_Click()
CommonDialog1.ShowColor
RichTextBox1.SelColor = CommonDialog1.Color
End Sub

Private Sub Form_Load()
Dim a As Integer
For a = 1 To Screen.FontCount
cmbFont.AddItem Screen.Fonts(a)
Next a
For a = 6 To 72 Step 2
cmbFSize.AddItem a
Next a
'MsgBox "Welcome to Siva's Text Editor", vbOKOnly, "Siva"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackColor = &H8000000F
Command2.BackColor = &H8000000F

End Sub

Private Sub Form_Resize()
On Error Resume Next
RichTextBox1.Height = Me.Height - 1330
RichTextBox1.Width = Me.Width - 200
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = True
If Saved = False Then
Dim MessageBox As String
MessageBox = MsgBox("Do yo want to save the unsaved document?", vbYesNoCancel, "Siva")
    If MessageBox = vbYes Then
    CommonDialog1.ShowSave
    RichTextBox1.SaveFile CommonDialog1.FileName
    Timer1.Enabled = True
    Dialog.Show
    End If
    If MessageBox = vbNo Then
    Timer1.Enabled = True
    Dialog.Show
    End If
Else
    If Saved = True Then
    Timer1.Enabled = True
    Dialog.Show
    End If
End If
End Sub

Private Sub Format_Click()

CommonDialog1.ShowFont

RichTextBox1.SelBold = CommonDialog1.FontBold
RichTextBox1.SelItalic = CommonDialog1.FontItalic
RichTextBox1.SelFontSize = CommonDialog1.FontSize
RichTextBox1.SelFontName = CommonDialog1.FontName
RichTextBox1.SelStrikeThru = CommonDialog1.FontStrikethru
RichTextBox1.SelUnderline = CommonDialog1.FontUnderline
'RichTextBox1.SelText = CommonDialog1.Font

End Sub

Private Sub InsertDateAndTime_Click()
RichTextBox1.SelText = Now
End Sub

Private Sub InsertDateDate_Click()
RichTextBox1.SelText = Date
End Sub

Private Sub InsertDateTime_Click()
RichTextBox1.SelText = Time
End Sub

Private Sub Italic_Click()
RichTextBox1.SelItalic = Not RichTextBox1.SelItalic
End Sub

Private Sub New_Click()
RichTextBox1.Text = ""
End Sub

Private Sub Open_Click()
Command2_Click
End Sub

Private Sub Save_Click()
On Error Resume Next

If Saved = False Then
    CommonDialog1.ShowSave
    RichTextBox1.SaveFile CommonDialog1.FileName
    Saved = True
    SaveLocation = CommonDialog1.FileName
Else
    If Saved = True Then
    RichTextBox1.SaveFile SaveLocation
    End If
    End If
End Sub

Private Sub SaveAs_Click()
On Error Resume Next
CommonDialog1.FileName = ""
CommonDialog1.ShowSave

RichTextBox1.SaveFile CommonDialog1.FileName
End Sub

Private Sub SelectAll_Click()
RichTextBox1.SelStart = 0
RichTextBox1.SelLength = Len(RichTextBox1.Text)
'RichTextBox1.SelText = RichTextBox1.Text
End Sub

Private Sub Space2_Click()

End Sub

Private Sub Time_Click()
RichTextBox1.SelText = Time

End Sub

Private Sub Timer1_Timer()
Me.WindowState = 0
If Me.Height > 700 Then
Me.Height = Me.Height - 70
Else
End
End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
If Button.Index = 1 Then
RichTextBox1.Text = ""
End If

If Button.Index = 2 Then
Command2_Click
End If



If Button.Index = 3 Then
  On Error Resume Next

    If Saved = False Then
        CommonDialog1.ShowSave
        RichTextBox1.SaveFile CommonDialog1.FileName
        Saved = True
        SaveLocation = CommonDialog1.FileName
        Else
        If Saved = True Then
        RichTextBox1.SaveFile SaveLocation
        End If
    End If
End If





If Button.Index = 5 Then
RichTextBox1.SelBold = Not RichTextBox1.SelBold
End If

If Button.Index = 6 Then
RichTextBox1.SelItalic = Not RichTextBox1.SelItalic
End If

If Button.Index = 7 Then
RichTextBox1.SelUnderline = Not RichTextBox1.SelUnderline
End If

If Button.Index = 8 Then
RichTextBox1.SelBold = False
RichTextBox1.SelItalic = False
RichTextBox1.SelUnderline = False
End If

If Button.Index = 9 Then
CommonDialog1.ShowColor
RichTextBox1.SelColor = CommonDialog1.Color
End If



If Button.Index = 10 Then
CommonDialog1.ShowFont
RichTextBox1.SelBold = CommonDialog1.FontBold
RichTextBox1.SelItalic = CommonDialog1.FontItalic
RichTextBox1.SelFontSize = CommonDialog1.FontSize
RichTextBox1.SelFontName = CommonDialog1.FontName
RichTextBox1.SelStrikeThru = CommonDialog1.FontStrikethru
RichTextBox1.SelUnderline = CommonDialog1.FontUnderline
'RichTextBox1.SelText = CommonDialog1.Font
End If


If Button.Index = 12 Then
Clipboard.SetText RichTextBox1.SelText
RichTextBox1.SelText = ""
End If

If Button.Index = 13 Then
Clipboard.SetText RichTextBox1.SelText
End If

If Button.Index = 14 Then
RichTextBox1.SelText = Clipboard.GetText
End If

If Button.Index = 17 Then
RichTextBox1.SelFontSize = RichTextBox1.SelFontSize - 1
End If

If Button.Index = 16 Then
RichTextBox1.SelFontSize = RichTextBox1.SelFontSize + 1
End If

If Button.Index = 16 Then
RichTextBox1.SelFontSize = RichTextBox1.SelFontSize + 1
End If


If Button.Index = 19 Then
Form2.Show
End If

If Button.Index = 21 Then
PopupMenu InsertDate
'RichTextBox1.SelText = Now
End If

If Button.Index = 23 Then
RichTextBox1.SelCharOffset = -72
End If

If Button.Index = 24 Then
RichTextBox1.SelCharOffset = 20
End If

If Button.Index = 26 Then
RichTextBox1.SelText = LCase(RichTextBox1.SelText)
End If

If Button.Index = 27 Then
RichTextBox1.SelText = UCase(RichTextBox1.SelText)
End If


End Sub


Private Sub Underline_Click()
RichTextBox1.SelUnderline = Not RichTextBox1.SelUnderline
End Sub
