Attribute VB_Name = "Module2"
'Just paste the whole enchilada into ONE (1) new module in
'your project and POW. You have the ability to CALL ANY
'launchable program you can think of from your project.
'Call it like this:

'Call Shell("Whatever.txt") 'if the Whatever is in the same
'directory folder as the .EXE that you compile.

'or

'Call Shell("C:\wherever\it\is.htm")

Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" _
Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As _
String, ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'#############################################################
'# This code was written by Emmett Dixson (c)1999. You may alter
'# this code, trade, steal, borrow, lend or give away this code.
'# However, this code has been regisered with the Library of
'# Congress as a literary acheivement and as such excludes it
'# from being known or proclaimed as "PUBLIC DOMAIN".
'#---------------You may NOT remove this header---------------
'#------------------You may NOT SELL this work----------------
'#----YES! You MAY use this work for commercial purposes------
'#---This code MAY NOT be sold or redistributed for profit----
'#-------- I wish you every success in your projects ---------
'#------------------------ Visit me at -----------------------
'#------------------http://developer.ecorp.net ---------------
'#-----------------FREE Visual Basic Source Code -------------
'##############################################################

'For best results paste everything into a NEW MODULE and be sure
'you SAVE the module to your project.

'Works for Win3.x, Win95,Win98,WinNT and EVEN Win2000(don't ask!)

'Don't change anything...just paste everything into ONE
'MODULE that you can add to a project.
            
Function Shell(Program As String, Optional ShowCmd As Long = _
vbNormalNoFocus, Optional ByVal WorkDir As Variant) As Long

    Dim FirstSpace As Integer, Slash As Integer

    If Left(Program, 1) = """" Then
        FirstSpace = InStr(2, Program, """")


        If FirstSpace <> 0 Then
            Program = Mid(Program, 2, FirstSpace - 2) & _
              Mid(Program, FirstSpace + 1)
            FirstSpace = FirstSpace - 1
        End If

    Else
        FirstSpace = InStr(Program, " ")
    End If

    If FirstSpace = 0 Then FirstSpace = Len(Program) + 1

    If IsMissing(WorkDir) Then

        For Slash = FirstSpace - 1 To 1 Step -1
            If Mid(Program, Slash, 1) = "\" Then Exit For
        Next

        If Slash = 0 Then
            WorkDir = CurDir
        ElseIf Slash = 1 Or Mid(Program, Slash - 1, 1) = ":" Then
            WorkDir = Left(Program, Slash)
        Else
            WorkDir = Left(Program, Slash - 1)
        End If

    End If

    Shell = ShellExecute(0, vbNullString, _
    Left(Program, FirstSpace - 1), LTrim(Mid(Program, _
    FirstSpace)), WorkDir, ShowCmd)
    'If Shell < 32 Then VBA.Shell Program, ShowCmd 'To raise Error

End Function


