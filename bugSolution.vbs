Function factorial(x)
  Dim result
  If x > 0 Then
    result = x * factorial(x - 1)
  Else
    result = 1
  End If
  factorial = result
End Function
MsgBox factorial(5)