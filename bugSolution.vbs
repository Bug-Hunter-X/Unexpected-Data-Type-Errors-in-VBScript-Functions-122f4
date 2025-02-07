Function f(a, b)
  If IsEmpty(a) Or IsEmpty(b) Then
    Err.Raise vbError, , "Both parameters must be provided."
    Exit Function 'Important: Exit after raising the error
  End If
  
  'Explicit type checking for robustness
  If VarType(a) <> vbString Or VarType(b) <> vbString Then
    Err.Raise vbError, , "Both parameters must be strings."
    Exit Function
  End If

  ' ... rest of the function using a and b
End Function