Attribute VB_Name = "Module1"
'e.g, 1024= one thousand twenty four,
 'from range 0 to 999 trillion. And since you'v got the code
 'you can modify it for more by adding few lines, like if
 'you know what comes after trillion you can easily add it
 'by just adding one line. You can use this code in your
 'programs but do mention the name of the author.
 'If you want to contact the author
 'my e-mail address is solo_sevensix@yahoo.com
 'Dont forget to visit my site-> http://www.solosoftware.co.nr
 '                                       ---Author Solomon Manalo
 Option Explicit
 
 '////Add This Code to a CLASS MODULE   \{.cls\}
 
 Private Data(9, 3) As String
 
 Private Sub Class_Initialize()
     'Data for conversion
     Data(0, 0) = "one": Data(1, 0) = "two": Data(2, 0) = "three"
     Data(3, 0) = "four": Data(4, 0) = "five": Data(5, 0) = "six"
     Data(6, 0) = "seven": Data(7, 0) = "eight": Data(8, 0) = "nine"
     Data(9, 0) = "ten"
     Data(0, 1) = "hundred": Data(1, 1) = "ten": Data(2, 1) = "twenty"
     Data(3, 1) = "thirty": Data(4, 1) = "fourty": Data(5, 1) = "fifty"
     Data(6, 1) = "sixty": Data(7, 1) = "seventy": Data(8, 1) = "eighty"
     Data(9, 1) = "ninety"
     Data(0, 3) = "ten": Data(1, 3) = "eleven": Data(2, 3) = "twelve"
     Data(3, 3) = "thirteen": Data(4, 3) = "fourteen": Data(5, 3) = "fifteen"
     Data(6, 3) = "sixteen": Data(7, 3) = "seventeen": Data(8, 3) = "eighteen"
     Data(9, 3) = "nineteen"
 
 End Sub
 
 Public Function ToWords(ByVal NumberStr As String) As String
     Dim z As String, X As String, Temp As String, c As String
     Dim a As Integer, b As Integer, i As Integer
     Dim iPos As Integer
     
     'remove redundant spaces
     NumberStr = Trim(Replace(NumberStr, ",", ""))
     a = Len(NumberStr)
     Temp = NumberStr
     If Val(NumberStr) = 0 Then
         ToWords = "zero!"
         Exit Function
     End If
     
     'get rid of any decimals
     iPos = InStr(Temp, ".")
     If iPos > 0 Then Temp = Left(Temp, iPos - 1)
     
     
     While ((a Mod 3) <> 0)
         Temp = "0" & Temp
         a = Len(Temp)
     Wend
     NumberStr = Temp
     For i = a - 2 To 1 Step -3
         b = b + 1
         Temp = Mid(NumberStr, i, 3)
         z = ""
         '  "Intelligent" routines
         '------------------------
         If Temp <> "000" Then
             c = Left(Temp, 1)
             If c <> "0" Then z = " " & Data(Val(c) - 1, 0) & " hundred"
             c = Mid(Temp, 2, 1)
             If c <> "0" Then
                 If c <> "1" Then
                     z = z & " " & Data(Val(c), 1)
                 Else
                     z = z & " " & Data(Val(Right(Temp, 2)) - 10, 3)
                 End If
             End If
             If Right(Temp, 1) <> "0" And Mid(Temp, 2, 1) <> "1" Then z = z & " " & Data(Val(Right(Temp, 1)) - 1, 0)
         End If
         '------------------------
         If z <> "" Then
             Select Case b
                 Case 1:
                     X = z
                 Case 2:
                     X = z & " thousand" & X
                 Case 3:
                     X = z & " million" & X
                 Case 4:
                     X = z & " billion" & X
                 Case 5:
                     X = z & " trillion" & X
                 Case Else:
                     Exit Function
                 'you can easily add more range
                 'like Case 6: can be "zillion"? :) (whatever)
             End Select
         End If
     Next
     ToWords = X
     
 
 End Function
 
 Private Function Replace(ByVal sInput As String, _
    sFind As String, sReplace As String) As String
  
 'USED HERE INSTEAD OF BUILT-IN REPLACE FUNCTION
 'SO THAT CLASS WILL WORK WITH VB5
 Dim lPos As Long
 Dim sAns As String
 Dim sWkg As String
 
 sAns = ""
 sWkg = sInput
 
 lPos = InStr(sWkg, sFind)
 
 If lPos <> 0 Then
     Do
         If lPos >= Len(sWkg) Then
            sAns = sAns & Left(sWkg, Len(sWkg) - 1) & sReplace
         Else
             sAns = sAns & Left(sWkg, lPos - 1) & sReplace
        End If
         sWkg = Mid(sWkg, lPos + 1)
         lPos = InStr(sWkg, sFind)
         DoEvents
     Loop While lPos > 0
     sAns = sAns & sWkg
 Else
     sAns = sInput
 End If
 Replace = sAns
 End Function

