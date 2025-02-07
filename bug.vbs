Function f(a, b)
  If IsEmpty(a) Or IsEmpty(b) Then
    Err.Raise vbError, , "Both parameters must be provided."
  End If
  ' ... rest of the function
End Function