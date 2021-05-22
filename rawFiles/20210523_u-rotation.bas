Sub Macro1()
'
' Macro1 Macro
'

'
    
    'アクティブシートの最終行を取得
    '
    For i = 0 To 130
        Dim raceName, raceData, raceData2, raceData3
    
        'レース名
        '
        raceName = Range("A474").Select
        Selection.Copy
        
        Range("O1000000").End(xlUp).Offset(1).Select
        ActiveSheet.Paste
        
        'レースデータ
        '
        raceData = Split(Cells(480, 2), " ")
        raceData2 = Split(raceData(2), "ｍ（")
        raceData3 = Split(raceData2(1), "）")
        
        'バ場、地面、m、距離、回り
        '
        Range("T1000000").End(xlUp).Offset(1).Value = raceData(0)
        Range("Z1000000").End(xlUp).Offset(1).Value = raceData(1)
        Range("P1000000").End(xlUp).Offset(1).Value = raceData2(0)
        Range("X1000000").End(xlUp).Offset(1).Value = raceData3(0)
        Range("V1000000").End(xlUp).Offset(1).Value = raceData3(1)
        
        '時期
        '
        Range("B482").Select
        Selection.Copy
        Range("N1000000").End(xlUp).Offset(1).Select
        ActiveSheet.Paste
        
        
        
        '行削除
        '
        Rows(474).Delete
        Rows(474).Delete
        Rows(474).Delete
        Rows(474).Delete
        Rows(474).Delete
        Rows(474).Delete
        Rows(474).Delete
        Rows(474).Delete
        Rows(474).Delete
        Rows(474).Delete
    Next i
End Sub
