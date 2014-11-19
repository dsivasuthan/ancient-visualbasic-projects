VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "My Phone Book"
   ClientHeight    =   2250
   ClientLeft      =   765
   ClientTop       =   615
   ClientWidth     =   4095
   Icon            =   "PhonebookForm1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4095
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3720
      Top             =   0
   End
   Begin VB.CommandButton Command6 
      Caption         =   ">>"
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
      Left            =   2040
      TabIndex        =   9
      ToolTipText     =   "Next"
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "<<"
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
      TabIndex        =   8
      ToolTipText     =   "Previous"
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete"
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
      Left            =   2760
      TabIndex        =   6
      ToolTipText     =   "Click to delete the current record"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add New"
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
      TabIndex        =   5
      ToolTipText     =   "Click to add a new record"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
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
      Left            =   1440
      TabIndex        =   4
      ToolTipText     =   "Click to save the current record entered"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtPhone 
      DataField       =   "Phone"
      DataSource      =   "adoPhone"
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
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   3
      ToolTipText     =   "Enter the telephone no"
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox txtName 
      DataField       =   "Name"
      DataSource      =   "adoPhone"
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
      Left            =   1440
      TabIndex        =   2
      ToolTipText     =   "Enter the title to the telephone no"
      Top             =   120
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc adoPhone 
      Height          =   330
      Left            =   0
      Top             =   2880
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=Files\Phone2VB.mdb"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=Files\Phone2VB.mdb"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Phone"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Form1 
      Alignment       =   2  'Center
      Caption         =   "Dhayalan Sivasuthan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      ToolTipText     =   "The Author"
      Top             =   1965
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   "Tel No"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Add 
         Caption         =   "Add new record"
      End
      Begin VB.Menu Save 
         Caption         =   "Save the current record"
      End
      Begin VB.Menu Close 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu Author 
      Caption         =   "About the Author"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Add_Click()
Command2_Click
End Sub

Private Sub Author_Click()
Dialog.Show
Me.Enabled = False
End Sub

Private Sub Close_Click()
Dim Ex As String
Ex = MsgBox("Do you really want to exit?", vbYesNo, "Siva")
If Ex = vbYes Then
Timer1.Enabled = True
Else
Cancel = True
End If
End Sub

Private Sub Command1_Click()
On Error Resume Next
adoPhone.Recordset.Fields("Name") = txtName.Text
adoPhone.Recordset.Fields("Phone") = txtPhone.Text
adoPhone.Recordset.Update
Command6_Click
End Sub

Private Sub Command2_Click()
'On Error Resume Next
adoPhone.Recordset.AddNew
txtName.SetFocus
End Sub

Private Sub Command3_Click()
On Error Resume Next
Confirm = MsgBox("Are you sure you want to delete this record?", vbYesNo, "Deletion Confirmation")
If Confirm = vbYes Then
adoPhone.Recordset.Delete
MsgBox "Record Deleted!", , "Message"
txtName = ""
txtPhone = ""
Else
MsgBox "Record Not Deleted!", , "Message"
End If

End Sub

Private Sub Command4_Click()
On Error Resume Next
txtName = ""
txtPhone = ""
txtName.SetFocus
End Sub

Private Sub Command5_Click()
On Error Resume Next
If Not adoPhone.Recordset.BOF Then
adoPhone.Recordset.MovePrevious
If adoPhone.Recordset.BOF Then
adoPhone.Recordset.MoveNext
End If
End If

End Sub

Private Sub Command6_Click()
On Error Resume Next
If Not adoPhone.Recordset.EOF Then
adoPhone.Recordset.MoveNext
If adoPhone.Recordset.EOF Then
adoPhone.Recordset.MovePrevious
End If
End If

End Sub

Private Sub Form_Load()
'MsgBox "Welcome to Siva's Phone Book", , "Siva"
End Sub


Private Sub Form_Unload(Cancel As Integer)
Cancel = True
'Dim Ex As String
'Ex = MsgBox("Do you really want to exit?", vbYesNo, "Siva")
'If Ex = vbYes Then
Timer1.Enabled = True
Dialog.Show
'End If
End Sub

Private Sub Save_Click()
Command1_Click
End Sub

Private Sub Timer1_Timer()
If Me.Height > 678 Then
Me.Height = Me.Height - 50
Else
End
End If
End Sub
