Sub Macro1()
  Dim arr, arr2
  arr = Split("a,b,c,d", ",")
  arr2 = Split(Cells(2, 3), " ")
  
  Range("D1").Value = arr(0)
  Range("D2").Value = arr2(0)
  
End Sub
