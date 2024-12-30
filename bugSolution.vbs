Function MyFunc(param1)
  If IsEmpty(param1) Then
    Exit Function
  End If
  'Rest of the function code
  'Example: 
  Result = param1 * 2
  MyFunc = Result
End Function