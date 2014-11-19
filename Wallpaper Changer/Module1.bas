Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function SystemParametersInfo Lib "user32" _
Alias "SystemParametersInfoA" (ByVal uAction As Long, _
ByVal uParam As Long, ByVal lpvParam As String, _
ByVal fuWinIni As Long) As Long

Public DEx As Boolean



