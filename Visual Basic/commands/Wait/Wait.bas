Sub Wait(Duration)
  Dim numTime
  numTime = Timer
  Do While Timer - numTime < Duration
    DoEvents
  Loop
End Sub