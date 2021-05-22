Sub Macro1()
  
  Range("A100000").End(xlUp).Offset(1).Select
  Selection.Copy
  
  Range("B1").Select
  ActiveSheet.Paste
  
End Sub
