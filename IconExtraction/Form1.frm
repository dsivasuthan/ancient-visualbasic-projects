VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Icon Extractor - Dhayalan Sivasuthan"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12390
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   12390
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4560
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "DLL Files|*.dll"
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Shut Down computer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   4560
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   6360
      Width           =   10455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      TabIndex        =   1
      Top             =   6360
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   240
      ScaleHeight     =   5955
      ScaleWidth      =   11955
      TabIndex        =   0
      Top             =   240
      Width           =   12015
      Begin VB.CommandButton Command3 
         Caption         =   "Save as"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   5520
         Visible         =   0   'False
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
'Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
'Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Const EWX_SHUTDOWN As Long = 1
Private Declare Function ExitWindowsEx Lib "user32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long

Sub Shutdown_Computer()
    Dim lngResult As Long
    lngResult = ExitWindowsEx(EWX_SHUTDOWN, 0&)
End Sub


Private Sub Extract_Icon(FileName As String)
  On Error Resume Next

  Dim i As Long
  Dim lIcon As Long
  Dim Xok As Long
  Dim Yok As Long
  Yok = 0
  Xok = 0
  Do
    lIcon = ExtractIcon(App.hInstance, Text1.Text, i)
    If lIcon = 0 Then Exit Do
    i = i + 1
    DestroyIcon lIcon
     'DestroyIcon lIcon
   'Picture1.Cls
   lIcon = ExtractIcon(App.hInstance, Text1.Text, i)       ' i = icon index
   Picture1.AutoSize = True
   Picture1.AutoRedraw = True
   DrawIcon Picture1.hdc, Xok, Yok, lIcon
   Xok = Xok + 35
 If Xok > Picture1.Width - 11250 Then
 'Yok = 32
 Yok = Yok + 35
 Xok = 0
 End If
   
   Picture1.Refresh
'MsgBox "Click ok to show next icon"
  Loop
  Command3.Visible = True
  If i = 0 Then
    MsgBox "No Icons in this file!"
Command3.Visible = False
  End If
End Sub


Private Sub Command1_Click()
CommonDialog1.Filter = "DLL Files|*.dll"
CommonDialog1.ShowOpen
If CommonDialog1.FileName = "" Then Exit Sub
Text1.Text = CommonDialog1.FileName

Extract_Icon Text1.Text
End Sub

Private Sub Command2_Click()
Shutdown_Computer
End Sub

Private Sub Command3_Click()
CommonDialog1.Filter = "Bitmap File|*.bmp"
CommonDialog1.FileName = ""
CommonDialog1.ShowSave
If CommonDialog1.FileName = "" Then Exit Sub
SavePicture Picture1.Image, CommonDialog1.FileName


End Sub
