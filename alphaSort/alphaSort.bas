Public Function alphaSort(toSort As String) As String
'Creates an alphanumeric string which when sorted upon 
' will produce human expected results.

'e.g. 
'Consider the following series ( a1, a001, a10, a010, a0100 )
'A standard sort will produce( a001, a010, a0100, a1, a10 )
'This function will produce a string which when used for a sort 
' will yield ( a1, a001, a10, a010, a0100 )

'How it works: Each time a numeric digit is encountered the loop 
' executes to count the number of digits which follow. 
' A two digit prefix is then inserted before the series of digits
' Leading zeros are handled the same way but appended as a suffix

  Dim sLZTemp As String 'Leading Zero temp
  Dim iLZCount As Integer 'Leading Zero counter
  Dim sNumTemp As String 'Number temp
  Dim iNumCount As Integer 'Digit counter
  Dim sOut as string 'Temp string output 
  On Error goto LEH
  While Len(toSort)
    'is the first character a number?
    If IsNumeric(Left(toSort, 1)) Then
      While IsNumeric(Left(toSort, 1))
        If Left(toSort, 1) = "0" And iNumCount = 0 Then
          iLZCount = iLZCount + 1 'Leading Zero counter
        Else
          iNumCount = iNumCount + 1 'significant digit counter
          sNumTemp = sNumTemp & Left(toSort, 1) 'remember all significant digits
        End If
        toSort = Mid(toSort, 2) 'loop on next character
      Wend
      'Leading Zeros counted, significant digits counted...
      sLZTemp = sLZTemp & Format(iLZCount, "00") 'leading zero sort code for later...
      sNumTemp = Format(iNumCount, "00") & sNumTemp
      sOut = sOut & sNumTemp
      iNumCount = 0 'reset for next imbedded digit string...
      iLZCount = 0
      sNumTemp = ""
    End If
    'next character not a number, loop on next character in string...
    sOut = sOut & Left(toSort, 1)
    toSort = Mid(toSort, 2)
  Wend
  sOut = sOut & " " & sLZTemp 'concatenate leading zero sort code
  alphaSort = sOut 'set output 
  exit function
  LEH: 'Local error handler
  alphaSort ="Error"
End Functiona

