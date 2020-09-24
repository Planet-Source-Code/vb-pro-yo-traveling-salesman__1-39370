Attribute VB_Name = "Module1"
Dim corrx(1 To 10)
 Dim corry(1 To 10)
 Dim fastest As Long
 Dim fastestset
  Sub beginprog()
  fastest = 2000000000
    Randomize
  For i = 1 To 10
  corrx(i) = Int(Rnd * 1500)
  corry(i) = Int(Rnd * 1500)
  Next i
  Form1.List3.AddItem corrx(1) & "   " & corry(1)
  Form1.List3.AddItem corrx(2) & "   " & corry(2)
  Form1.List3.AddItem corrx(3) & "   " & corry(3)
  Form1.List3.AddItem corrx(4) & "   " & corry(4)
  Form1.List3.AddItem corrx(5) & "   " & corry(5)
  End Sub
  Sub thebest()
  Form1.Text2.Text = fastestset
  Form1.Text3.Text = fastest
  End Sub
  Sub checkfast(a, b, c, d, e, f)
  aa = a
  xone = corrx(a)
  xtwo = corrx(b)
  xthree = corrx(c)
  xfour = corrx(d)
  xfive = corrx(e)
  xsix = corrx(f)
  yone = corry(a)
  ytwo = corry(b)
  ythree = corry(c)
  yfour = corry(d)
  yfive = corry(e)
  ysix = corry(f)
  xtotal = xtwo - xone
  ytotal = ytwo - yone
Form1.Shape1.Left = corrx(1)
Form1.Shape1.Top = corry(1)
Form1.Shape2.Left = corrx(2)
Form1.Shape2.Top = corry(2)
Form1.Shape3.Left = corrx(3)
Form1.Shape3.Top = corry(3)
Form1.Shape4.Left = corrx(4)
Form1.Shape4.Top = corry(4)
Form1.Shape5.Left = corrx(5)
Form1.Shape5.Top = corry(5)
Form1.Shape6.Left = corrx(6)
Form1.Shape6.Top = corry(6)
  xtotal = xtotal * xtotal
  ytotal = ytotal * ytotal
  finaltotal = xtotal + ytotal
  finaltotalone = Square_Root_Of(finaltotal)
'part two
  xtotal = xthree - xtwo
  ytotal = ythree - ytwo
  xtotal = xtotal * xtotal
  ytotal = ytotal * ytotal
  finaltotal = xtotal + ytotal
  finaltotaltwo = Square_Root_Of(finaltotal)
  'part three
  xtotal = xfour - xthree
  ytotal = yfour - ythree
  xtotal = xtotal * xtotal
  ytotal = ytotal * ytotal
  finaltotal = xtotal + ytotal
  finaltotalthree = Square_Root_Of(finaltotal)
  'part four
  xtotal = xfive - xfour
  ytotal = yfive - yfour
  xtotal = xtotal * xtotal
  ytotal = ytotal * ytotal
  finaltotal = xtotal + ytotal
  finaltotalfour = Square_Root_Of(finaltotal)
  'part five
  xtotal = xsix - xfive
  ytotal = ysix - yfive
  xtotal = xtotal * xtotal
  ytotal = ytotal * ytotal
  finaltotal = xtotal + ytotal
  finaltotalfive = Square_Root_Of(finaltotal)
  'part six
    xtotal = xsix - xone
  ytotal = ysix - yone
  xtotal = xtotal * xtotal
  ytotal = ytotal * ytotal
  finaltotal = xtotal + ytotal
  finaltotalsix = Square_Root_Of(finaltotal)
  'final part
   thr = InStr(finaltotalone, ".") 'find where to strip out more data
    finaltotalone = Left(finaltotalone, thr - 1)
     thr = InStr(finaltotaltwo, ".") 'find where to strip out more data
    finaltotaltwo = Left(finaltotaltwo, thr - 1)
     thr = InStr(finaltotalthree, ".") 'find where to strip out more data
   finaltotalthree = Left(finaltotalthree, thr - 1)
     thr = InStr(finaltotalfour, ".") 'find where to strip out more data
    finaltotalfour = Left(finaltotalfour, thr - 1)
    thr = InStr(finaltotalfive, ".") 'find where to strip out more data
    finaltotalfive = Left(finaltotalfive, thr - 1)
    thr = InStr(finaltotalsix, ".") 'find where to strip out more data
    finaltotalsix = Left(finaltotalsix, thr - 1)
    Dim totalforset As Long
    Form1.Text1.Text = finaltotalone
    totalforset = finaltotalone
    Form1.Text1.Text = finaltotaltwo
    totalforset = totalforset + finaltotaltwo
    Form1.Text1.Text = finaltotalthree
    totalforset = totalforset + finaltotalthree
    Form1.Text1.Text = finaltotalfour
    totalforset = totalforset + finaltotalfour
    Form1.Text1.Text = finaltotalfive
    totalforset = totalforset + finaltotalfive
    Form1.Text1.Text = finaltotalsix
    totalforset = totalforset + finaltotalsix
    If totalforset < fastest Then
    fastest = totalforset
    fastestset = a & " " & b & " " & c & " " & d & "  " & e & "  " & f
    End If
    Form1.List2.AddItem a & " " & b & " " & c & " " & d & "  " & e & "  " & f
    Form1.List1.AddItem totalforset
    End Sub
  Public Function Square_Root_Of(ArgX)
'AUTHOR:   Jay Tanner
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=18499&lngWId=1
' V1.3
  Dim X As Variant      ' Argument - Positive or negative value
  
  Dim a As Variant      ' Any general approximation to square root
  Dim b As Variant      ' Next successive approximation to square root
  
  Dim k  As Integer     ' Cycle loop control counter
  Dim i  As String      ' Represents the square root of minus 1
  
' Check for invalid numeric argument
  X = Trim(ArgX): If IsNumeric(X) = False Then GoTo ERROR_HANDLER

  X = CDec(X)  ' Convert argument into decimal data type
  
' Account for a negative argument
  i = "": If X < 0 Then X = -X: i = " i"
  
' Check for zero argument
  If X = 0 Then Square_Root_Of = 0: Exit Function

' Use VB square root as 1st approximation
  a = Sqr(X)
  
'   Very primitive loop to grind out the square root using a series
'   of successive approximations, starting with (A).
    k = 50 ' Set limit of cycles to 50 max - More than enough.
CYCLE:
    b = (a + X / a) / 2 ' Compute next approx (B) from (A)
    
' Check if finished
  If (b = a) Or k <= 0 Then GoTo DONE

' Rinse, lather, repeat until done
  a = b        ' Update approx to current value
  k = k - 1    ' Update limit counter
  GoTo CYCLE
DONE:
  Square_Root_Of = Trim(b & i)
  Exit Function
  
ERROR_HANDLER:
  Square_Root_Of = "ERROR: Invalid numeric argument"
  End Function
