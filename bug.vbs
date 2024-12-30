Function MyFunc(param1)
  If IsEmpty(param1) Then
    ' Handle empty parameter
    Exit Function 'This line was missing, causing an error when param1 is empty
  End If
  'Rest of the function code
End Function