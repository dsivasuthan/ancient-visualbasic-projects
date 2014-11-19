Attribute VB_Name = "Module1"
' ----------------------------
' Constants & API Declarations
' ----------------------------

Option Explicit

Global lIcon&
Global sSourcePgm$
Global sDestFile$

Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long


' ----------------------------
' Constants & API Declarations
' ----------------------------

Private Const EWX_SHUTDOWN As Long = 1
Private Declare Function ExitWindowsEx Lib "user32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long


' ---------
' Function
' ---------

Sub Shutdown_Computer()
    Dim lngResult As Long
    lngResult = ExitWindowsEx(EWX_SHUTDOWN, 0&)
End Sub
