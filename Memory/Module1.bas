Attribute VB_Name = "Module1"
Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)


Public Type MEMORYSTATUS
    dwlength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPage As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Public DEx As Boolean
Public Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2

Public Sub StayOnTop(frm As Form)
  SetWindowPos frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub


