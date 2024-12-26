Option Explicit

' Function to demonstrate safe string to number conversion
Function SafeConvertToInt(strNum)
  Dim intNum
  On Error Resume Next
  intNum = CInt(strNum)
  If Err.Number <> 0 Then
    Err.Clear
    intNum = 0 ' Or handle the error as appropriate
  End If
  SafeConvertToInt = intNum
End Function

' Example of early binding
Dim objFSO As Object
Set objFSO = CreateObject("Scripting.FileSystemObject")
' ... use objFSO ...
Set objFSO = Nothing

' Example demonstrating explicit type checking
Dim strValue, numValue
strValue = "123"
If IsNumeric(strValue) Then
  numValue = CInt(strValue)
  ' ... process numValue ...
Else
  ' ... handle the case where strValue is not a number ...
End If