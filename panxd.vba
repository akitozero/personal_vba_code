'需要添加参数，参数：sheet名字 要定位单元格的内容
'返回单元格的行数和列数
Function find_range(searchName As String, searchRange As String, searchSheet As String) As Range
    Dim wk_book As Workbook
    Dim wk_sheet As Worksheet
    Dim rng As Range
    Dim column_count As Integer
    Dim StrFind As String
    
    Set wk_book = Application.Workbooks("VBA练习表")
    Set wk_sheet = wk_book.Worksheets(searchSheet)
    
    'Debug.Print "你输入的内容是" & findName
    StrFind = searchName
    With wk_sheet.Range(searchRange)
        Set rng = .Find(What:=StrFind, _
        After:=.Cells(.Cells.Count), _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False)
        If Not rng Is Nothing Then
            Application.Goto rng, True
        End If
    End With
    Set find_range = rng
End Function
'参数：要复制值所在的sheet，目的复制对象的sheet，要复制值的纵坐标，目的复制对象的纵坐标
Function copy_val(ByVal x As Byte, ByVal y As Byte) As Byte()

End Function
Sub outputSub(str As String)
    Debug.Print "你输入的内容是：" & str
End Sub
Sub test_panxd()
    Dim wk_book As Workbook
    Dim wk_sheet As Worksheet
    Dim wk_sheet2 As Worksheet
    Set wk_book = Application.Workbooks("VBA练习表")
    Set wk_sheet = wk_book.Worksheets("日排行")
    Set wk_sheet2 = wk_book.Worksheets("总排行")
    
    Application.DisplayAlerts = False
    wk_sheet.Range("A1").CurrentRegion.Copy wk_sheet2.Range("B3")
    Application.DisplayAlerts = True
    
    Debug.Print "test2"
End Sub
'复制指定单元格给内容但另一张表指定单元格里
Sub test_panxd2()
    Dim rng1 As Range
    Dim rng2 As Range
    
    Set rng1 = Sheets("日排行").Range("B3")
    Set rng2 = Sheets("总排行").Range("B3")
    rng1.Copy rng2
End Sub
Sub test_panxd3()
    Dim rng1 As Range
    Dim rng2 As Range
    
'    Set rng1 = Sheets("日排行").Range("A1")
'    Set rng2 = Sheets("总排行").Range("A1")
    Set rng1 = Sheets("日排行").Cells(1, 1)
    Set rng2 = Sheets("总排行").Cells(1, 1)
    rng2.Value = rng2.Value + rng1.Value
    rng1.Value = 0
End Sub
Function sync_value_to_range(userName As String, task As String) As Integer
    Dim srcRng As Range
    Dim dstRng As Range
    Dim rng1 As Range
    Dim rng2 As Range
    Dim srcRow As Integer
    Dim srcColumn As Integer
    
    Set srcRng = find_range(userName, "B:B", "日排行")
    srcRow = srcRng.Row
    Set dstRng = find_range(userName, "B:B", "日排行")
    dstRow = dstRng.Row
    
    Set srcRng = find_range(task, "2:2", "日排行")
    srcColumn = srcRng.Column
    Set dstRng = find_range(task, "2:2", "总排行")
    dstColumn = dstRng.Column
    
'    Debug.Print "11 srcRow" & srcRow
'    Debug.Print "11 srcColumn" & srcColumn
    Set rng1 = Sheets("日排行").Cells(srcRow, srcColumn)
    Set rng2 = Sheets("总排行").Cells(dstRow, dstColumn)
    rng2.Value = rng2.Value + rng1.Value
    rng1.Value = 0
End Function
Sub test_find_range()
    Dim srcRng As Range
    Dim dstRng As Range
    Dim rng1 As Range
    Dim rng2 As Range
    Dim srcRow As Integer
    Dim srcColumn As Integer
    Dim i As Integer
    
    Dim rangeTask(1 To 10) As String
    rangeTask(1) = "任务"
    rangeTask(2) = "阅读"
    rangeTask(3) = "日记"
    rangeTask(4) = "复述"
    rangeTask(5) = "10个问题"
    rangeTask(6) = "回答问题"
    rangeTask(7) = "主题"
    rangeTask(8) = "成就"
    rangeTask(9) = "看书"
    rangeTask(10) = "复习anki"
    
    'find_range
'    Set srcRng = find_range("潘兴俤", "B:B", "日排行")
'    srcRow = srcRng.Row
'
'    Set dstRng = find_range("潘兴俤", "B:B", "日排行")
'    dstRow = dstRng.Row
'
'    Set srcRng = find_range("任务", "2:2", "总排行")
'    srcColumn = srcRng.Column
'
'    Set dstRng = find_range("任务", "2:2", "总排行")
'    dstColumn = dstRng.Column
'
'    Set rng1 = Sheets("日排行").Cells(srcRow, srcColumn)
'    Set rng2 = Sheets("总排行").Cells(dstRow, dstColumn)
'    rng2.Value = rng2.Value + rng1.Value
'    rng1.Value = 0
'    Debug.Print "rng1.value" & rng1.Value
'    Debug.Print "rng2.value" & rng2.Value
    For i = 1 To 10
        sync_value_to_range "潘兴俤", rangeTask(i)
    Next
    'Range(A1, A1).Select
    
    'Set rng1 = find_range("潘兴俤", "B:B", "日排行")
    'sync_value_to_range "潘兴俤", "任务"
    'sync_value_to_range "潘兴俤", "阅读"
    
'    Debug.Print "6行数是" & srcRow
'    Debug.Print "6列数是" & srcColumn
    
    'rng = find_range()
    '根据行数和列数选中指定单元格
    'Cells(srcRow, srcColumn).Value = 122
    
    'rng.Select
    
End Sub
Sub 测试用()
   'MsgBox "hehe"
   '测试sub过程传参
'    Dim str As String
'    str = "潘兴俤"
'    outputSub str
    'test_panxd3
    'test_find_range
    'test_panxd3
    test_find_range
End Sub

