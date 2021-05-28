Sub MACRO1()
    '現時点だと動かないマクロ
    '
    'ストアド名取得
    '
    
    Dim preObj, Obj, TableName
    Dim rowNum
    Dim cellObj, cellTable
    Dim nowCell
    cellObj = 1
    cellTable = 5
    rowNum = 5
    preObj = Cells(rowNum, cellObj).Value
    nowCell = 5
    Debug.Print preObj
    
    For i = 0 To 130
        
        Obj = Cells(rowNum, cellObj).Value
        Debug.Print Obj
        'ストアド名が前行と同じ
        '
        
        If preObj = Obj Then
            Cells(nowCell, 6).Value = Cells(nowCell, 6).Value + " " + Cells(rowNum, cellTable).Value
        Else
            nowCell = nowCell + 1
        End If
        
        rowNum = rowNum + 1
    Next i

 

End Sub
